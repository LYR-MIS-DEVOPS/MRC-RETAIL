VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form CashFlowMgrOld 
   Caption         =   "Manager's Cash Flow"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLyrAmt 
      Height          =   285
      Left            =   4920
      TabIndex        =   85
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLyrCv 
      Height          =   285
      Left            =   3720
      TabIndex        =   84
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCashOutAmt 
      Height          =   285
      Left            =   2880
      TabIndex        =   83
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCashOutCv 
      Height          =   285
      Left            =   2280
      TabIndex        =   82
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtExpAmt 
      Height          =   285
      Left            =   1680
      TabIndex        =   81
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtExpCv 
      Height          =   285
      Left            =   1200
      TabIndex        =   80
      Top             =   9720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtMgrCFId 
      Height          =   285
      Left            =   480
      TabIndex        =   79
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid flxOR 
      Height          =   735
      Left            =   8520
      TabIndex        =   74
      Top             =   9360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxDC 
      Height          =   735
      Left            =   6720
      TabIndex        =   73
      Top             =   9360
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxRC 
      Height          =   735
      Left            =   4440
      TabIndex        =   72
      Top             =   9360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   0
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      Caption         =   "New Cash Flow"
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
      Image           =   "CashFlowManager.frx":0000
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPV 
      Height          =   615
      Left            =   1680
      TabIndex        =   65
      Top             =   9360
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   18
      Tag             =   "XYF"
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Breakdown"
      TabPicture(0)   =   "CashFlowManager.frx":059A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdRefresh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Collection and Expenses"
      TabPicture(1)   =   "CashFlowManager.frx":05B6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxCF"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Preview"
      TabPicture(2)   =   "CashFlowManager.frx":05D2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CRViewer1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Attachment"
      TabPicture(3)   =   "CashFlowManager.frx":05EE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CRViewer2"
      Tab(3).Control(1)=   "cmdPrint"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   5760
         TabIndex        =   93
         Tag             =   "XYF"
         Top             =   6320
         Width           =   5535
         Begin VB.TextBox txtPrepBy 
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   103
            Tag             =   "XYF"
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtAprvBy 
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   102
            Tag             =   "XYF"
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtCVamt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Left            =   3960
            MaxLength       =   20
            TabIndex        =   98
            Tag             =   "XYF"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtCVno 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            MaxLength       =   20
            TabIndex        =   97
            Tag             =   "XYF"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtCVchk 
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   95
            Tag             =   "XYF"
            Top             =   200
            Width           =   3495
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   105
            Tag             =   "XYF"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label5 
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
            Left            =   120
            TabIndex        =   104
            Tag             =   "XYF"
            Top             =   1350
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "CV AMT:"
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
            Left            =   2280
            TabIndex        =   99
            Tag             =   "XYF"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "CV NUMBER:"
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
            TabIndex        =   96
            Tag             =   "XYF"
            Top             =   630
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "CV CHECK DETAILS:"
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
            TabIndex        =   94
            Tag             =   "XYF"
            Top             =   225
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   240
         TabIndex        =   89
         Tag             =   "XYF"
         Top             =   6320
         Width           =   5415
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
            Height          =   1155
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   100
            Tag             =   "XYF"
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox txtNotedBy 
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
            Left            =   1440
            TabIndex        =   90
            Tag             =   "XYF"
            Top             =   1920
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.Label Label7 
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
            Index           =   0
            Left            =   480
            TabIndex        =   101
            Tag             =   "XYF"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Noted By:"
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
            TabIndex        =   91
            Tag             =   "XYF"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   240
         TabIndex        =   86
         Tag             =   "XYF"
         Top             =   360
         Width           =   3855
         Begin MSComCtl2.DTPicker CFdate 
            Height          =   315
            Left            =   1200
            TabIndex        =   87
            Tag             =   "XYF"
            Top             =   180
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMMM dd, yyyy"
            Format          =   64028675
            CurrentDate     =   38900
         End
         Begin VB.Label Label2 
            Caption         =   "Select Date:"
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
            TabIndex        =   88
            Tag             =   "XYF"
            Top             =   210
            Width           =   1815
         End
      End
      Begin CRVIEWERLibCtl.CRViewer CRViewer2 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   76
         Tag             =   "xy"
         Top             =   840
         Width           =   11175
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   0   'False
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin lvButton.lvButtons_H cmdPrint 
         Height          =   375
         Left            =   -74880
         TabIndex        =   75
         Tag             =   "XYF"
         Top             =   435
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Update Attachment"
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
         Image           =   "CashFlowManager.frx":060A
         cBack           =   -2147483633
      End
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   7695
         Left            =   -74880
         TabIndex        =   64
         Tag             =   "xy"
         Top             =   480
         Width           =   11175
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   0   'False
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   0   'False
         EnableRefreshButton=   0   'False
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   0   'False
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin VB.Frame Frame3 
         Caption         =   "C H E C K S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   240
         TabIndex        =   54
         Tag             =   "XYF"
         Top             =   960
         Width           =   11055
         Begin VB.TextBox txtChk 
            Height          =   285
            Left            =   120
            TabIndex        =   63
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid flxChk 
            Height          =   2055
            Left            =   120
            TabIndex        =   55
            Tag             =   "XY"
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   3625
            _Version        =   393216
            Rows            =   1
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   315
            BackColorBkg    =   -2147483633
            ScrollBars      =   2
            BorderStyle     =   0
            FormatString    =   $"CashFlowManager.frx":0BA4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "C O I N S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2925
         Left            =   5760
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   3360
         Width           =   5535
         Begin VB.TextBox txtTotalCoins 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            TabIndex        =   56
            Tag             =   "XY"
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Tag             =   "XY"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Tag             =   "XY"
            Top             =   555
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   2
            Left            =   360
            TabIndex        =   9
            Tag             =   "XY"
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   3
            Left            =   360
            TabIndex        =   10
            Tag             =   "XY"
            Top             =   1185
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   4
            Left            =   360
            TabIndex        =   11
            Tag             =   "XY"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   5
            Left            =   360
            TabIndex        =   12
            Tag             =   "XY"
            Top             =   1815
            Width           =   1215
         End
         Begin VB.Label lblCoinTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3720
            TabIndex        =   62
            Tag             =   "XY"
            Top             =   1815
            Width           =   1575
         End
         Begin VB.Label lblCoinTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3720
            TabIndex        =   61
            Tag             =   "XY"
            Top             =   1500
            Width           =   1575
         End
         Begin VB.Label lblCoinTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3720
            TabIndex        =   60
            Tag             =   "XY"
            Top             =   1185
            Width           =   1575
         End
         Begin VB.Label lblCoinTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3720
            TabIndex        =   59
            Tag             =   "XY"
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label lblCoinTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3720
            TabIndex        =   58
            Tag             =   "XY"
            Top             =   555
            Width           =   1575
         End
         Begin VB.Label lblCoinTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   3720
            TabIndex        =   57
            Tag             =   "XY"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   1920
            TabIndex        =   53
            Tag             =   "XY"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   1920
            TabIndex        =   52
            Tag             =   "XY"
            Top             =   555
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   1920
            TabIndex        =   51
            Tag             =   "XY"
            Top             =   870
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   1920
            TabIndex        =   50
            Tag             =   "XY"
            Top             =   1185
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   1920
            TabIndex        =   49
            Tag             =   "XY"
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   1920
            TabIndex        =   48
            Tag             =   "XY"
            Top             =   1815
            Width           =   375
         End
         Begin VB.Label lblCoin 
            Alignment       =   1  'Right Justify
            Caption         =   "10.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   47
            Tag             =   "XY"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblCoin 
            Alignment       =   1  'Right Justify
            Caption         =   "5.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   46
            Tag             =   "XY"
            Top             =   555
            Width           =   975
         End
         Begin VB.Label lblCoin 
            Alignment       =   1  'Right Justify
            Caption         =   "1.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   2280
            TabIndex        =   45
            Tag             =   "XY"
            Top             =   870
            Width           =   975
         End
         Begin VB.Label lblCoin 
            Alignment       =   1  'Right Justify
            Caption         =   "0.25"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   2280
            TabIndex        =   44
            Tag             =   "XY"
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label lblCoin 
            Alignment       =   1  'Right Justify
            Caption         =   "0.10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   2280
            TabIndex        =   43
            Tag             =   "XY"
            Top             =   1500
            Width           =   975
         End
         Begin VB.Label lblCoin 
            Alignment       =   1  'Right Justify
            Caption         =   "0.05"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   2280
            TabIndex        =   42
            Tag             =   "XY"
            Top             =   1815
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "B I L L S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2925
         Left            =   240
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   3360
         Width           =   5415
         Begin VB.TextBox txtTotalBills 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3480
            TabIndex        =   13
            Tag             =   "XY"
            Top             =   2475
            Width           =   1575
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   6
            Left            =   480
            TabIndex        =   6
            Tag             =   "XY"
            Top             =   2130
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   5
            Left            =   480
            TabIndex        =   5
            Tag             =   "XY"
            Top             =   1815
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   4
            Left            =   480
            TabIndex        =   4
            Tag             =   "XY"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   3
            Left            =   480
            TabIndex        =   3
            Tag             =   "XY"
            Top             =   1185
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   2
            Left            =   480
            TabIndex        =   2
            Tag             =   "XY"
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   1
            Left            =   480
            TabIndex        =   1
            Tag             =   "XY"
            Top             =   555
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   0
            Left            =   480
            TabIndex        =   0
            Tag             =   "XY"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblBillTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   3480
            TabIndex        =   40
            Tag             =   "XY"
            Top             =   2130
            Width           =   1575
         End
         Begin VB.Label lblBillTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3480
            TabIndex        =   39
            Tag             =   "XY"
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblBillTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3480
            TabIndex        =   38
            Tag             =   "XY"
            Top             =   1500
            Width           =   1575
         End
         Begin VB.Label lblBillTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3480
            TabIndex        =   37
            Tag             =   "XY"
            Top             =   1185
            Width           =   1575
         End
         Begin VB.Label lblBillTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3480
            TabIndex        =   36
            Tag             =   "XY"
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label lblBillTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3480
            TabIndex        =   35
            Tag             =   "XY"
            Top             =   555
            Width           =   1575
         End
         Begin VB.Label lblBillTtl 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   3480
            TabIndex        =   34
            Tag             =   "XY"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblBill 
            Alignment       =   1  'Right Justify
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   2160
            TabIndex        =   33
            Tag             =   "XY"
            Top             =   2130
            Width           =   975
         End
         Begin VB.Label lblBill 
            Alignment       =   1  'Right Justify
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   2160
            TabIndex        =   32
            Tag             =   "XY"
            Top             =   1815
            Width           =   975
         End
         Begin VB.Label lblBill 
            Alignment       =   1  'Right Justify
            Caption         =   "50"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   2160
            TabIndex        =   31
            Tag             =   "XY"
            Top             =   1500
            Width           =   975
         End
         Begin VB.Label lblBill 
            Alignment       =   1  'Right Justify
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   2160
            TabIndex        =   30
            Tag             =   "XY"
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblBill 
            Alignment       =   1  'Right Justify
            Caption         =   "200"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   2160
            TabIndex        =   29
            Tag             =   "XY"
            Top             =   870
            Width           =   975
         End
         Begin VB.Label lblBill 
            Alignment       =   1  'Right Justify
            Caption         =   "500"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   28
            Tag             =   "XY"
            Top             =   555
            Width           =   975
         End
         Begin VB.Label lblBill 
            Alignment       =   1  'Right Justify
            Caption         =   "1000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   27
            Tag             =   "XY"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   26
            Tag             =   "XY"
            Top             =   2130
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   1800
            TabIndex        =   25
            Tag             =   "XY"
            Top             =   1815
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   1800
            TabIndex        =   24
            Tag             =   "XY"
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1800
            TabIndex        =   23
            Tag             =   "XY"
            Top             =   1185
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1800
            TabIndex        =   22
            Tag             =   "XY"
            Top             =   870
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   21
            Tag             =   "XY"
            Top             =   555
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1800
            TabIndex        =   20
            Tag             =   "XY"
            Top             =   240
            Width           =   375
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxCF 
         Height          =   7695
         Left            =   -74880
         TabIndex        =   78
         Tag             =   "XY"
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   13573
         _Version        =   393216
         Rows            =   0
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   14737632
         ScrollBars      =   2
         MergeCells      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   9720
         TabIndex        =   92
         Tag             =   "XFY"
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         Caption         =   "&Refresh"
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
         Image           =   "CashFlowManager.frx":0C49
         cBack           =   -2147483633
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5520
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowManager.frx":11E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowManager.frx":177D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowManager.frx":1D17
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowManager.frx":22B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowManager.frx":284B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowManager.frx":2DE5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   2
      Left            =   3600
      TabIndex        =   67
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      Caption         =   "View Cash Flow"
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
      Image           =   "CashFlowManager.frx":337F
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   3
      Left            =   5400
      TabIndex        =   68
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      Caption         =   "Print Preview"
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
      Image           =   "CashFlowManager.frx":3919
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   4
      Left            =   7200
      TabIndex        =   69
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      Caption         =   "Quit Cash Flow"
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
      Image           =   "CashFlowManager.frx":3EB3
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   5
      Left            =   9000
      TabIndex        =   70
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      Caption         =   "Save Cash Flow"
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
      Image           =   "CashFlowManager.frx":444D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   6
      Left            =   10800
      TabIndex        =   71
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      Caption         =   "Cancel Cash Flow"
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
      Image           =   "CashFlowManager.frx":49E7
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   1
      Left            =   1800
      TabIndex        =   77
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      Caption         =   "Edit Cash Flow"
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
      Image           =   "CashFlowManager.frx":4F81
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxRemarks 
      Height          =   735
      Left            =   2400
      TabIndex        =   106
      Top             =   10200
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   1440
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11805
      m_MDI_Height    =   9765
      OriginalBorder  =   2
      OriginalCaption =   27
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MANAGER'S CASH FLOW"
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
      Index           =   5
      Left            =   8280
      TabIndex        =   17
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MANAGER'S CASH FLOW"
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
      Index           =   4
      Left            =   8310
      TabIndex        =   16
      Top             =   510
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC"
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
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC"
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
      Index           =   2
      Left            =   270
      TabIndex        =   14
      Top             =   510
      Width           =   2415
   End
   Begin VB.Menu mnuChecks 
      Caption         =   "Checks"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "CashFlowMgrOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mcLyr As CRAXDRT.Application
Dim mcReport As CRAXDRT.Report
Dim rsRPT As New ADODB.Recordset
Dim ActRow As Integer, ActCol As Integer
Dim ChkLstRow As Integer
Dim NewCF As Boolean

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    Frame2.Enabled = Not Bln
    Frame3.Enabled = Not Bln
    Frame4.Enabled = Not Bln
    cmdRefresh.Enabled = Not Bln
End Sub

Private Sub ClearEntries()
    Dim i As Integer
    
    CFdate.Value = Now
    txtPrepBy.Text = ""
    txtAprvBy.Text = ""
    flxChk.Rows = 1
    flxChk.Rows = 2
    flxCF.Rows = 0
    For i = 0 To 5
        txtBill(i).Text = ""
        txtCoin(i).Text = ""
    Next i
    txtBill(6).Text = ""
    txtTotalBills.Text = ""
    txtTotalCoins.Text = ""
    
End Sub

Private Sub tbSaveCancel(Bln As Boolean)
    Dim x As Integer
    
    If Bln Then
        For x = 0 To 4
            cmdTrans(x).Visible = False
        Next x
        cmdTrans(5).Left = 0
        cmdTrans(6).Left = cmdTrans(5).Width
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For x = 0 To 4
            cmdTrans(x).Visible = True
        Next x
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    End If
End Sub

Private Sub PrintCF()
    Dim txtObj As TextObject
    Dim x As Integer, myTop As Integer
    Dim FontCourier As Boolean, NoLine As Boolean
    Set mcLyr = New CRAXDRT.Application
    Set mcReport = mcLyr.NewReport
    
    mcReport.PaperSize = crPaperLetter
    mcReport.TopMargin = 0
    mcReport.BottomMargin = 0
    mcReport.Sections(1).Height = 0
    mcReport.Sections(2).Height = 0
    With mcReport.Sections(3)
        .Height = 18000
        Set txtObj = .AddTextObject("BINOS MARKETING", 0, 300)
        txtObj.Width = 12000: txtObj.HorAlignment = crHorCenterAlign: txtObj.Font.Name = "Arial": txtObj.Font.Bold = True: txtObj.Font.Size = 12: txtObj.Height = 350
        Set txtObj = .AddTextObject("L & L Bldg., Sobrecary St., Tagum City", 0, 650)
        txtObj.Width = 12000: txtObj.HorAlignment = crHorCenterAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 300
        Set txtObj = .AddTextObject("CASHIER'S CASH FLOW", 0, 1600)
        txtObj.Width = 12000: txtObj.HorAlignment = crHorCenterAlign: txtObj.Font.Name = "Arial": txtObj.Font.Bold = True: txtObj.Font.Size = 12: txtObj.Height = 350
        Set txtObj = .AddTextObject("Date : " & Format$(CFdate.Value, "Mmm. dd, yyyy"), 9000, 1950)
        txtObj.Width = 2500: txtObj.HorAlignment = crLeftAlign: txtObj.Font.Name = "Arial": txtObj.Font.Underline = True: txtObj.Font.Size = 10: txtObj.Height = 250
        myTop = 2300: FontCourier = False
        For x = 0 To flxCF.Rows - 5
            If flxCF.TextMatrix(x, 0) <> "" Then
                Set txtObj = .AddTextObject(flxCF.TextMatrix(x, 0), 500, myTop)
                txtObj.Width = 3000: txtObj.HorAlignment = crLeftAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Font.Bold = True: txtObj.Height = 300
                If flxCF.TextMatrix(x, 3) <> "" Then
                    Set txtObj = .AddTextObject(flxCF.TextMatrix(x, 3) & "   ", 9800, myTop)
                    txtObj.Width = 1600: txtObj.HorAlignment = crRightAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Font.Bold = True: txtObj.Height = 220: txtObj.BottomLineStyle = crLSSingleLine
                End If
                myTop = myTop + 250
                If flxCF.TextMatrix(x, 0) = "BILLS" Then FontCourier = True
                If flxCF.TextMatrix(x, 0) = "Total" Then FontCourier = False
                If flxCF.TextMatrix(x, 0) = "REMARKS:" Then NoLine = True Else NoLine = False
            ElseIf flxCF.TextMatrix(x, 1) <> "" Then
                Set txtObj = .AddTextObject(flxCF.TextMatrix(x, 1) & IIf(FontCourier, String$(30, "."), ""), 1500, myTop)
                txtObj.Width = 6000: txtObj.HorAlignment = crLeftAlign:  txtObj.Font.Size = 10:  txtObj.Height = 220
                If FontCourier Then txtObj.Font.Name = "Courier New" Else txtObj.Font.Name = "Arial"
                Set txtObj = .AddTextObject(IIf(flxCF.TextMatrix(x, 2) = "", " ", flxCF.TextMatrix(x, 2)) & "   ", 7800, myTop)
                txtObj.Width = 1700: txtObj.HorAlignment = crRightAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10:  txtObj.Height = 220:  txtObj.BottomLineStyle = IIf(NoLine, crLSNoLine, crLSSingleLine)
                If flxCF.TextMatrix(x, 3) <> "" Then
                    Set txtObj = .AddTextObject(flxCF.TextMatrix(x, 3) & "   ", 9800, myTop)
                    txtObj.Width = 1600: txtObj.HorAlignment = crRightAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Font.Bold = True: txtObj.Height = 220: txtObj.BottomLineStyle = crLSSingleLine
                End If
                myTop = myTop + 250
            Else
                myTop = myTop + 250
            End If
        Next x
        Set txtObj = .AddTextObject("Prepared By:", 500, myTop)
        txtObj.Width = 3000: txtObj.HorAlignment = crLeftAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 220
        Set txtObj = .AddTextObject("Approved By:", 7800, myTop)
        txtObj.Width = 3000: txtObj.HorAlignment = crLeftAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 220
        myTop = myTop + 750
        Set txtObj = .AddTextObject(txtPrepBy.Text, 1500, myTop)
        txtObj.Width = 2300: txtObj.HorAlignment = crHorCenterAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 220: txtObj.Font.Underline = True
        Set txtObj = .AddTextObject(txtAprvBy.Text, 8500, myTop)
        txtObj.Width = 2300: txtObj.HorAlignment = crHorCenterAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 220: txtObj.Font.Underline = True
        myTop = myTop + 250
        Set txtObj = .AddTextObject("Cashier", 1500, myTop)
        txtObj.Width = 2300: txtObj.HorAlignment = crHorCenterAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 220
        Set txtObj = .AddTextObject("Manager", 8500, myTop)
        txtObj.Width = 2300: txtObj.HorAlignment = crHorCenterAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 220
        myTop = myTop + 350
        Set txtObj = .AddTextObject("Date Printed: " & Now, 1500, myTop)
        txtObj.Width = 8000: txtObj.HorAlignment = crLeftAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 8: txtObj.Height = 200
    End With
    mcReport.Sections(4).Height = 0
    mcReport.Sections(5).Height = 0
        
 '   CRViewer1.ReportSource = mcReport
    mcReport.PrintOut False, 1, , 1, 1
End Sub


Private Sub ComputeBills()
    Dim x As Integer
    Dim Ttl As Single
    
    Ttl = 0
    For x = 0 To 6
        Ttl = Ttl + PureNum(lblBillTtl(x).Caption)
    Next x
    txtTotalBills.Text = Format$(Ttl, "##,###,##0.00")

End Sub

Private Sub ComputeCoins()
    Dim x As Integer
    Dim Ttl As Single
    
    Ttl = 0
    For x = 0 To 5
        Ttl = Ttl + PureNum(lblCoinTtl(x).Caption)
    Next x
    txtTotalCoins.Text = Format$(Ttl, "##,###,##0.00")

End Sub

Function GetMgrCF(myField As String) As Variant
    Dim rsMgrCF As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT CASHFLOWMGR.* From CashFlowMgr " & _
          "WHERE (((CASHFLOWMGR.MGRCFID)=" & txtMgrCFId.Text & "));"
    'mySQL
    Qry = "SELECT CASHFLOWMGR.* From CashFlowMgr " & _
          "WHERE CASHFLOWMGR.MGRCFID=" & txtMgrCFId.Text & ";"
     
    With rsMgrCF
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If myField = "ExpenseCV" Then
                GetMgrCF = IIf(IsNull(.Fields(myField)), "", Format(.Fields(myField), "0000"))
            Else
                GetMgrCF = IIf(IsNull(.Fields(myField)), "", .Fields(myField))
            End If
        Else
            GetMgrCF = ""
        End If
        .Close
    End With
End Function

Function OnlineSalesRetAmt() As Currency
    Dim rsOL As New ADODB.Recordset
    Dim Qry As String, MyInv As String
    Dim TtlRet As Currency
    
    TtlRet = 0
    'Qry = "SELECT PAYMENTS.INVOICENO From PAYMENTS " & _
          "WHERE (((PAYMENTS.SPAREPARTS)>0) AND ((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((PAYMENTS.REMARKS) Like 'ON-LINE PAYMENT%'));"
    'mySQL
    Qry = "SELECT PAYMENTS.INVOICENO From PAYMENTS " & _
          "WHERE PAYMENTS.SPAREPARTS>0 AND PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND (PAYMENTS.REMARKS) Like 'ON-LINE PAYMENT%';"
    With rsOL
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                MyInv = !InvoiceNO
                'Qry = "SELECT SRETURNSP.TOTALAMT From SRETURNSP " & _
                      "WHERE (((SRETURNSP.INVOICE)='" & MyInv & "') AND ((SRETURNSP.PROCESS)='REPLACE'));"
                'mySQL
                Qry = "SELECT SRETURNSP.TOTALAMT From SRETURNSP " & _
                      "WHERE SRETURNSP.INVOICE='" & MyInv & "' AND SRETURNSP.PROCESS='REPLACE';"
                TtlRet = TtlRet + Val(GetFieldValue(Qry))
                .MoveNext
            Loop
        End If
        .Close
    End With
    OnlineSalesRetAmt = TtlRet
End Function


Private Sub ProcessCashFlow()
    Dim PartsInv As String, OthersOR As String, ServiceOR As String, PartsOR As String, cfQry As String
    Dim i As Integer, BrkDwnRow As Integer
    Dim pdcTtl As Single, BrkDwnTtl As Currency, NetColl As Currency
    Dim CashierColl As Currency
    Dim SPcash As String, SPonAcct As String, DpOR As String, SrvcAmt As String, DPamt As String, OtherAmt As String, OLinv As String, OLamt As String
    
                  
    'SrvcAmt = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
              "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SERVICE))<>0));"
        
    'ServiceOR = "SELECT PAYMENTS.ORNO, PAYMENTS.SERVICE, PAYMENTS.PAYDATE From PAYMENTS " & _
                "Where ((Not (PAYMENTS.ORNO) Is Null) And ((PAYMENTS.SERVICE) <> 0) And ((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.ORNO;"
    'SPonAcct = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
               "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"

    'PartsOR = "SELECT PAYMENTS.ORNO, PAYMENTS.SPAREPARTS, PAYMENTS.PAYDATE From PAYMENTS " & _
              "Where ((Not (PAYMENTS.ORNO) Is Null) And ((PAYMENTS.SPAREPARTS) <> 0) And ((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.ORNO;"
    'SPcash = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
             "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.INVOICENO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
    'PartsInv = "SELECT PAYMENTS.INVOICENO, PAYMENTS.SPAREPARTS, PAYMENTS.PAYDATE From PAYMENTS " & _
               "WHERE ((Not (PAYMENTS.INVOICENO) Is Null) AND ((PAYMENTS.SPAREPARTS)<>0) AND ((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.INVOICENO;"

    'DpOR = "SELECT PAYMENTS.ORNO, PAYMENTS.DOWNPAYMENT, PAYMENTS.PAYDATE From PAYMENTS " & _
           "Where ((Not (PAYMENTS.ORNO) Is Null) And ((PAYMENTS.DOWNPAYMENT) <> 0) And ((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.ORNO;"
    'DPamt = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT From PAYMENTS " & _
            "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.DOWNPAYMENT))<>0));"

    'OthersOR = "SELECT PAYMENTS.ORNO, PAYMENTS.OTHERS, PAYMENTS.PAYDATE From PAYMENTS " & _
               "Where ((Not (PAYMENTS.ORNO) Is Null) And ((PAYMENTS.OTHERS) <> 0) And ((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.ORNO;"
    'OtherAmt = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
               "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
    'OLinv = "SELECT PAYMENTS.INVOICENO From PAYMENTS " & _
            "WHERE (((PAYMENTS.SPAREPARTS)>0) AND ((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((PAYMENTS.REMARKS)='ON-LINE PAYMENT'));"
    'OLamt = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
            "WHERE (((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((PAYMENTS.REMARKS)='ON-LINE PAYMENT'));"
    
    'mySQL
    SrvcAmt = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
              "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null) HAVING (Sum(PAYMENTS.SERVICE))<>0;"
        
    ServiceOR = "SELECT PAYMENTS.ORNO, PAYMENTS.SERVICE, PAYMENTS.PAYDATE From PAYMENTS " & _
                "Where (Not (PAYMENTS.ORNO) Is Null) And PAYMENTS.SERVICE <> 0 And PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.ORNO;"
    SPonAcct = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
               "Where PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null) HAVING (Sum(PAYMENTS.SPAREPARTS))<>0;"

    PartsOR = "SELECT PAYMENTS.ORNO, PAYMENTS.SPAREPARTS, PAYMENTS.PAYDATE From PAYMENTS " & _
              "Where (Not (PAYMENTS.ORNO) Is Null) And PAYMENTS.SPAREPARTS <> 0 And PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.ORNO;"
    SPcash = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
             "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.INVOICENO) Is Null) HAVING (Sum(PAYMENTS.SPAREPARTS))<>0;"
    PartsInv = "SELECT PAYMENTS.INVOICENO, PAYMENTS.SPAREPARTS, PAYMENTS.PAYDATE From PAYMENTS " & _
               "WHERE (Not (PAYMENTS.INVOICENO) Is Null) AND PAYMENTS.SPAREPARTS<>0 AND PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.INVOICENO;"

    DpOR = "SELECT PAYMENTS.ORNO, PAYMENTS.DOWNPAYMENT, PAYMENTS.PAYDATE From PAYMENTS " & _
           "Where (Not (PAYMENTS.ORNO) Is Null) And PAYMENTS.DOWNPAYMENT <> 0 And PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.ORNO;"
    DPamt = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT From PAYMENTS " & _
            "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null) HAVING (Sum(PAYMENTS.DOWNPAYMENT))<>0;"

    OthersOR = "SELECT PAYMENTS.ORNO, PAYMENTS.OTHERS, PAYMENTS.PAYDATE From PAYMENTS " & _
               "Where (Not (PAYMENTS.ORNO) Is Null) And PAYMENTS.OTHERS <> 0 And PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.ORNO;"
    OtherAmt = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
               "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null) HAVING (Sum(PAYMENTS.OTHERS))<>0;"
    OLinv = "SELECT PAYMENTS.INVOICENO From PAYMENTS " & _
            "WHERE PAYMENTS.SPAREPARTS>0 AND PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS='ON-LINE PAYMENT';"
    OLamt = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
            "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS='ON-LINE PAYMENT';"

    With flxCF
        .Rows = 0
        .AddItem "COLLECTION"
        .AddItem vbTab & "SPARE PARTS"
        .AddItem vbTab & Space$(5) & "CASH [INV#:" & GetOR(PartsInv) & vbTab & Format$(GetFieldValue(SPcash), "#,##0.00")
        .AddItem vbTab & Space$(5) & "ON-ACCOUNT [OR#:" & GetOR(PartsOR) & vbTab & Format$(GetFieldValue(SPonAcct), "#,##0.00")
        .AddItem vbTab & "SERVICE [OR#:" & GetOR(ServiceOR) & "]" & vbTab & Format$(GetFieldValue(SrvcAmt), "#,##0.00")
        .AddItem vbTab & "DOWN PAYMENT [OR#:" & GetOR(DpOR) & "]" & vbTab & Format$(GetFieldValue(DPamt), "#,##0.00")
        .AddItem vbTab & "Others [" & "OR#:" & GetOR(OthersOR) & "]" & vbTab & Format$(GetFieldValue(OtherAmt), "#,##0.00")
        .TextMatrix(6, 3) = FormatNumber(FlxTotal(2, 6, 2), 2)
        .AddItem "LESS"
        If NewCF Then
            txtExpCv.Text = CVNOS("EXPENSE")
            txtExpAmt.Text = ExpenseTotal
            txtCVno.Text = txtExpCv.Text
            txtCVamt.Text = Format(txtExpAmt.Text, "#,##0.00")
            txtCashOutCv.Text = CVNOS("TRANSFER")
            txtCashOutAmt.Text = GetCVamt("TRANSFER")
            txtLyrCv.Text = GetOR(OLinv)
            txtLyrAmt.Text = Val(GetFieldValue(OLamt)) - OnlineSalesRetAmt
        Else
            txtExpCv.Text = CVNOS("EXPENSE")
            txtExpAmt.Text = ExpenseTotal
            txtCVno.Text = GetMgrCF("ExpenseCV")
            txtCVamt.Text = Format(GetMgrCF("ExpenseAmt"), "#,##0.00")
            txtCashOutCv.Text = GetMgrCF("CashoutCV")
            txtCashOutAmt.Text = GetMgrCF("CashoutAmt")
            txtLyrCv.Text = GetOR(OLinv)
            txtLyrAmt.Text = Val(GetFieldValue(OLamt)) - OnlineSalesRetAmt
        End If
        .AddItem vbTab & "Expenses [CV# " & txtExpCv.Text & "]" & vbTab & Format$(txtExpAmt.Text, "#,##0.00")
        .AddItem vbTab & "Cash Out " & IIf(CVNOS("TRANSFER") <> "", "[CV# " & txtCashOutCv.Text & " TOF]", "") & vbTab & Format$(txtCashOutAmt.Text, "#,##0.00")
        .AddItem vbTab & "O.R. issued to Post Dated Checks " & ORforPDC & vbTab & Format$(AMTforPDC, "#,##0.00")
        'cfQry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
                "WHERE (((SRETURNSP.RETDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((SRETURNSP.PROCESS)='REPLACE') AND ((SPINVOICE.INVTYPE)='CASH'));"
        'mySQL
        cfQry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
                "WHERE SRETURNSP.RETDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND SRETURNSP.PROCESS='REPLACE' AND SPINVOICE.INVTYPE='CASH';"
        
        .AddItem vbTab & "SALES RETURN (Replace Invoice) [SR#:" & SalesReturnSeries & vbTab & Format$(GetFieldValue(cfQry), "#,##0.00")
        .AddItem vbTab & "On-line Payment " & IIf(txtLyrCv.Text <> "", "[INV# " & txtLyrCv.Text & "]", "") & vbTab & Format$(txtLyrAmt.Text, "#,##0.00")
        .TextMatrix(12, 3) = Format$(FlxTotal(8, 12, 2), "#,##0.00")
        nRow = 12
        NetColl = Val(Replace(.TextMatrix(6, 3), ",", "")) - Val(Replace(.TextMatrix(nRow, 3), ",", ""))
        .AddItem "NET COLLECTION" & vbTab & vbTab & vbTab & FormatNumber(NetColl, 2)
        .AddItem ""
        .AddItem "BREAKDOWN:"
        BrkDwnRow = .Rows
        .AddItem "CHECKS" '& vbTab & String$(15, "_") & vbTab & String$(15, "_")
        If flxChk.TextMatrix(flxChk.Rows - 1, 1) <> "" Or flxChk.TextMatrix(flxChk.Rows - 1, 3) <> "" Or flxChk.TextMatrix(flxChk.Rows - 1, 4) <> "" Or flxChk.TextMatrix(flxChk.Rows - 1, 6) <> "" Then
            For i = 1 To flxChk.Rows - 1
                .AddItem vbTab & "Check#" & flxChk.TextMatrix(i, 4) & ", " & flxChk.TextMatrix(i, 1) & ", " & flxChk.TextMatrix(i, 3) & vbTab & Format$(flxChk.TextMatrix(i, 6), "#,###,##0.00")
            Next i
            .TextMatrix(.Rows - 1, 3) = Format$(FlxTotal(BrkDwnRow, .Rows - 1, 2), "##,###,##0.00")
        End If
        .AddItem "" 'vbTab & String$(15, "_") & vbTab & String$(15, "_") & vbTab & String$(15, "_")
        .AddItem "BILLS"
        For i = 0 To 6
            .AddItem vbTab & Space$(6 - Len(txtBill(i).Text)) & txtBill(i).Text & Space$(5) & "X" & Space$(10 - Len(lblBill(i).Caption)) & lblBill(i).Caption & Space$(3) & String$(15, "..") & vbTab & lblBillTtl(i).Caption
        Next i
        .TextMatrix(.Rows - 1, 3) = txtTotalBills.Text
        .AddItem "COINS"
        For i = 0 To 5
            .AddItem vbTab & Space$(6 - Len(txtCoin(i).Text)) & txtCoin(i).Text & Space$(5) & "X" & Space$(10 - Len(lblCoin(i).Caption)) & lblCoin(i).Caption & Space$(3) & String$(15, "..") & vbTab & lblCoinTtl(i).Caption
        Next i
        .TextMatrix(.Rows - 1, 3) = txtTotalCoins.Text
        .AddItem "Total"
        BrkDwnTtl = FlxTotal(BrkDwnRow, .Rows - 1, 3)
        .TextMatrix(.Rows - 1, 3) = Format$(BrkDwnTtl, "#,##0.00")
        .AddItem vbTab & "Add CV Check: " & txtCVchk.Text & "(CV No. " & txtCVno.Text & ")" & vbTab & vbTab & txtCVamt.Text
        .AddItem vbTab & "TOTAL AMOUNT FOR DEPOSIT" & vbTab & vbTab & Format(BrkDwnTtl + Val(Replace(txtCVamt.Text, ",", "")), "#,##0.00")
        .AddItem ""
        .AddItem ""
       ' .AddItem "TOTAL" & vbTab & vbTab & vbTab & Format$(BrkDwnTtl + pdcTtl, "##,###,##0.00")
        .AddItem "REMARKS:"
        If Format(NetColl - BrkDwnTtl, "#,##0.00") <> 0# Then
            If NetColl - BrkDwnTtl < 0 Then
                .AddItem vbTab & "Excess : " & Format$(Abs(NetColl - BrkDwnTtl), "###,##0.00")
            ElseIf NetColl - BrkDwnTtl > 0 Then
                .AddItem vbTab & "SHORTAGE : " & Format$(NetColl - BrkDwnTtl, "###,##0.00")
            End If
        End If
        'If CountIBP > 0 Then Call DisplayIBPinFlx
        If CountReceiveChk > 0 Then Call DispRecChkinFlx
        If CountDepositedChk > 0 Then Call DispDepChkinFlx
        If AMTforPDC > 0 Then Call DispORChkinFlx
        If txtRemarks.Text <> "" Then .AddItem vbTab & txtRemarks.Text
        .AddItem ""
        .AddItem ""
        .AddItem ""
        .AddItem "Prepared By:" & vbTab & vbTab & "Approved By:"
        .AddItem ""
        .AddItem txtPrepBy.Text & vbTab & vbTab & txtAprvBy.Text
        .AddItem "   O.I.C." & vbTab & vbTab & "   Manager"
    End With
End Sub
Private Sub ProcessCF()
    Dim crCashier As New CFmanager
    Dim PartsInv As String, OthersOR As String, ServiceOR As String, PartsOR As String
    Dim i As Integer, BrkDwnRow As Integer
    Dim ChkSum As Single, Attch As Boolean
    Dim myRem As String, AQry As String
    Dim SPcash As String, SPonAcct As String, DpOR As String, SrvcAmt As String, DPamt As String, OtherAmt As String
    Dim cfQry As String, OLinv As String, OLamt As String
    
                  
    'SrvcAmt = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
              "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SERVICE))<>0));"
        
    'ServiceOR = "SELECT PAYMENTS.ORNO, PAYMENTS.SERVICE, PAYMENTS.PAYDATE From PAYMENTS " & _
                "Where ((Not (PAYMENTS.ORNO) Is Null) And ((PAYMENTS.SERVICE) <> 0) And ((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.ORNO;"
    'SPonAcct = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
               "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"

    'PartsOR = "SELECT PAYMENTS.ORNO, PAYMENTS.SPAREPARTS, PAYMENTS.PAYDATE From PAYMENTS " & _
              "Where ((Not (PAYMENTS.ORNO) Is Null) And ((PAYMENTS.SPAREPARTS) <> 0) And ((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.ORNO;"
    'SPcash = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
             "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.INVOICENO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
    'PartsInv = "SELECT PAYMENTS.INVOICENO, PAYMENTS.SPAREPARTS, PAYMENTS.PAYDATE From PAYMENTS " & _
               "WHERE ((Not (PAYMENTS.INVOICENO) Is Null) AND ((PAYMENTS.SPAREPARTS)<>0) AND ((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.INVOICENO;"

    'DpOR = "SELECT PAYMENTS.ORNO, PAYMENTS.DOWNPAYMENT, PAYMENTS.PAYDATE From PAYMENTS " & _
           "Where ((Not (PAYMENTS.ORNO) Is Null) And ((PAYMENTS.DOWNPAYMENT) <> 0) And ((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.ORNO;"
    'DPamt = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT From PAYMENTS " & _
            "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.DOWNPAYMENT))<>0));"

    'OthersOR = "SELECT PAYMENTS.ORNO, PAYMENTS.OTHERS, PAYMENTS.PAYDATE From PAYMENTS " & _
               "Where ((Not (PAYMENTS.ORNO) Is Null) And ((PAYMENTS.OTHERS) <> 0) And ((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#)) ORDER BY PAYMENTS.ORNO;"
    'OtherAmt = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
               "Where (((PAYMENTS.PAYDATE) = #" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
    'OLinv = "SELECT PAYMENTS.INVOICENO From PAYMENTS " & _
            "WHERE (((PAYMENTS.SPAREPARTS)>0) AND ((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((PAYMENTS.REMARKS)='ON-LINE PAYMENT'));"
    'OLamt = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
            "WHERE (((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((PAYMENTS.REMARKS)='ON-LINE PAYMENT'));"
    'mySQL
    SrvcAmt = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
              "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null) HAVING (Sum(PAYMENTS.SERVICE))<>0;"
        
    ServiceOR = "SELECT PAYMENTS.ORNO, PAYMENTS.SERVICE, PAYMENTS.PAYDATE From PAYMENTS " & _
                "Where (Not (PAYMENTS.ORNO) Is Null) And PAYMENTS.SERVICE <> 0 And PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.ORNO;"
    SPonAcct = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
               "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null) HAVING (Sum(PAYMENTS.SPAREPARTS))<>0;"

    PartsOR = "SELECT PAYMENTS.ORNO, PAYMENTS.SPAREPARTS, PAYMENTS.PAYDATE From PAYMENTS " & _
              "Where (Not (PAYMENTS.ORNO) Is Null) And PAYMENTS.SPAREPARTS <> 0 And PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.ORNO;"
    SPcash = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
             "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.INVOICENO) Is Null) HAVING (Sum(PAYMENTS.SPAREPARTS))<>0;"
    PartsInv = "SELECT PAYMENTS.INVOICENO, PAYMENTS.SPAREPARTS, PAYMENTS.PAYDATE From PAYMENTS " & _
               "WHERE (Not (PAYMENTS.INVOICENO) Is Null) AND PAYMENTS.SPAREPARTS<>0 AND PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.INVOICENO;"

    DpOR = "SELECT PAYMENTS.ORNO, PAYMENTS.DOWNPAYMENT, PAYMENTS.PAYDATE From PAYMENTS " & _
           "Where (Not (PAYMENTS.ORNO) Is Null) And PAYMENTS.DOWNPAYMENT <> 0 And PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.ORNO;"
    DPamt = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT From PAYMENTS " & _
            "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null) HAVING (Sum(PAYMENTS.DOWNPAYMENT))<>0;"

    OthersOR = "SELECT PAYMENTS.ORNO, PAYMENTS.OTHERS, PAYMENTS.PAYDATE From PAYMENTS " & _
               "Where (Not (PAYMENTS.ORNO) Is Null) And PAYMENTS.OTHERS <> 0 And PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' ORDER BY PAYMENTS.ORNO;"
    OtherAmt = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
               "Where PAYMENTS.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null) HAVING (Sum(PAYMENTS.OTHERS))<>0;"
    OLinv = "SELECT PAYMENTS.INVOICENO From PAYMENTS " & _
            "WHERE PAYMENTS.SPAREPARTS>0 AND PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS='ON-LINE PAYMENT';"
    OLamt = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
            "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS='ON-LINE PAYMENT';"
    
    With flxPV
        .Rows = 1
        .AddItem "[OR#:" & GetOR(ServiceOR) & "]" & vbTab & "str"
        .AddItem GetFieldValue(SrvcAmt)
        .AddItem "[INV#:" & GetOR(PartsInv) & "]" & vbTab & "str"
        .AddItem GetFieldValue(SPcash)
        .AddItem "[OR#:" & GetOR(OthersOR) & "]" & vbTab & "str"
        .AddItem GetFieldValue(OtherAmt)
        txtExpCv.Text = CVNOS("EXPENSE")
        txtExpAmt.Text = ExpenseTotal
        txtCashOutCv.Text = GetMgrCF("CashoutCV")
        txtCashOutAmt.Text = GetMgrCF("CashoutAmt")
        txtLyrCv.Text = GetOR(OLinv)
        txtLyrAmt.Text = Val(GetFieldValue(OLamt)) - OnlineSalesRetAmt
        .AddItem "[CV# " & txtExpCv.Text & "]" & vbTab & "str"
        .AddItem txtExpAmt.Text
        .AddItem IIf(txtCashOutCv.Text <> "", "[CV# " & txtCashOutCv.Text & " TOF]", "") & vbTab & "str"
        .AddItem txtCashOutAmt.Text
        .AddItem "[O.R.# " & ORforPDC & "]" & vbTab & "str"
        .AddItem AMTforPDC
        '.AddItem "" & vbTab & "str" 'Daily OR
        .AddItem SalesReturnAmt
        .AddItem IIf(txtLyrCv.Text <> "", "[INV# " & txtLyrCv.Text & "]", "") & vbTab & "str"
        .AddItem txtLyrAmt.Text 'Online Payment
    
        If flxChk.Rows = 2 Then
            ChkSum = 0
            If flxChk.TextMatrix(1, 6) <> "" Then
                .AddItem flxChk.TextMatrix(1, 4) & "-" & flxChk.TextMatrix(1, 2) & "-" & flxChk.TextMatrix(1, 3) & vbTab & "str"
                ChkSum = PureNum(flxChk.TextMatrix(1, 6))
                .AddItem PureNum(flxChk.TextMatrix(1, 6))
                .AddItem ChkSum
                For i = 1 To 4
                    .AddItem "" & vbTab & "str"
                    .AddItem 0
                    .AddItem 0
                Next i
            Else
                For i = 1 To 5
                    .AddItem "" & vbTab & "str"
                    .AddItem 0
                    .AddItem 0
                Next i
            End If
        End If
        If flxChk.Rows = 3 Then
            ChkSum = 0
            For i = 1 To 2
                .AddItem flxChk.TextMatrix(i, 4) & "-" & flxChk.TextMatrix(i, 2) & "-" & flxChk.TextMatrix(i, 3) & vbTab & "str"
                ChkSum = ChkSum + PureNum(flxChk.TextMatrix(i, 6))
                .AddItem PureNum(flxChk.TextMatrix(i, 6))
                If i = 1 Then
                    .AddItem 0
                Else
                    .AddItem ChkSum
                End If
            Next i
            For i = 1 To 3
                .AddItem "" & vbTab & "str"
                .AddItem 0
                .AddItem 0
            Next i
        End If
        If flxChk.Rows = 4 Then
            ChkSum = 0
            For i = 1 To 3
                .AddItem flxChk.TextMatrix(i, 4) & "-" & flxChk.TextMatrix(i, 2) & "-" & flxChk.TextMatrix(i, 3) & vbTab & "str"
                ChkSum = ChkSum + PureNum(flxChk.TextMatrix(i, 6))
                .AddItem PureNum(flxChk.TextMatrix(i, 6))
                If i < 3 Then
                    .AddItem 0
                Else
                    .AddItem ChkSum
                End If
            Next i
            For i = 1 To 2
                .AddItem "" & vbTab & "str"
                .AddItem 0
                .AddItem 0
            Next i
        End If
        
        If flxChk.Rows = 5 Then
            ChkSum = 0
            For i = 1 To 4
                .AddItem flxChk.TextMatrix(i, 4) & "-" & flxChk.TextMatrix(i, 2) & "-" & flxChk.TextMatrix(i, 3) & vbTab & "str"
                ChkSum = ChkSum + PureNum(flxChk.TextMatrix(i, 6))
                .AddItem PureNum(flxChk.TextMatrix(i, 6))
                If i < 4 Then
                    .AddItem 0
                Else
                    .AddItem ChkSum
                End If
            Next i
            .AddItem "" & vbTab & "str"
            .AddItem 0
            .AddItem 0
        End If
        If flxChk.Rows = 6 Then
            ChkSum = 0
            For i = 1 To 5
                .AddItem flxChk.TextMatrix(i, 4) & "-" & flxChk.TextMatrix(i, 2) & "-" & flxChk.TextMatrix(i, 3) & vbTab & "str"
                ChkSum = ChkSum + PureNum(flxChk.TextMatrix(i, 6))
                .AddItem PureNum(flxChk.TextMatrix(i, 6))
                If i < 5 Then
                    .AddItem 0
                Else
                    .AddItem ChkSum
                End If
            Next i
        End If
        If flxChk.Rows > 6 Then
            .AddItem "Total Amount (" & flxChk.Rows - 1 & " Checks - see attachment for details)" & vbTab & "str"
            ChkSum = 0
            For i = 1 To flxChk.Rows - 1
                ChkSum = ChkSum + PureNum(flxChk.TextMatrix(i, 6))
            Next i
            .AddItem ChkSum
            .AddItem ChkSum
            For i = 2 To 5
                .AddItem "" & vbTab & "str"
                .AddItem 0
                .AddItem 0
            Next i
        End If
        For i = 0 To 6
            .AddItem txtBill(i).Text
        Next i
        For i = 0 To 5
            .AddItem txtCoin(i).Text
        Next i
        .AddItem ChkSum
        'retrieve Inter Branch payment
        'If CountIBP > 0 Then myRem = IIf(myRem = "", DisplayIBP, myRem & vbCrLf & DisplayIBP)
        'retrieve receive checks
        flxRemarks.Rows = 0
        Attch = False
        If flxRC.Rows > 0 Then
            If (flxRC.Rows + flxDC.Rows + flxOR.Rows) >= 5 Then
                myRem = IIf(myRem <> "", myRem & vbCrLf, "")
                myRem = "Received Checks: " & flxRC.Rows & " Checks received (Please see attachment for details)"
                Attch = True
            Else
                myRem = IIf(myRem <> "", myRem & vbCrLf & "Checks Received:", "Checks Received:")
                For i = 0 To flxRC.Rows - 1
                    myRem = myRem & vbCrLf & Space$(5) & "PR#: " & flxRC.TextMatrix(i, 0) & ", PR Date:" & Format$(CFdate.Value, "mm/dd/yy") & ", Customer: " & flxRC.TextMatrix(i, 2) & ", Check#:" & flxRC.TextMatrix(i, 3) & ", Bank:" & flxRC.TextMatrix(i, 4) & ", Date:" & flxRC.TextMatrix(i, 5) & ", Amt: " & flxRC.TextMatrix(i, 6)
                    flxRemarks.AddItem IIf(i = 0, "Checks Received: ", Space$(5)) & "PR#: " & flxRC.TextMatrix(i, 0) & ", PR Date:" & Format$(CFdate.Value, "mm/dd/yy") & ", Customer: " & flxRC.TextMatrix(i, 2) & ", Check#:" & flxRC.TextMatrix(i, 3) & ", Bank:" & flxRC.TextMatrix(i, 4) & ", Date:" & flxRC.TextMatrix(i, 5) & ", Amt: " & flxRC.TextMatrix(i, 6)
                Next i
            End If
        End If
        'retrieve deposited checks
        If flxDC.Rows > 0 Then
            If (flxRC.Rows + flxDC.Rows + flxOR.Rows) >= 5 Then
                myRem = IIf(myRem <> "", myRem & vbCrLf, "")
                myRem = myRem & "Deposited Checks: " & flxDC.Rows & " Checks deposited (Please see attachment for details)"
                Attch = True
            Else
                myRem = IIf(myRem <> "", myRem & vbCrLf & "Deposited Checks:", "Deposited Checks:")
                For i = 0 To flxDC.Rows - 1
                    myRem = myRem & vbCrLf & Space$(5) & "PR#: " & flxDC.TextMatrix(i, 0) & ", PR Date:" & flxDC.TextMatrix(i, 1) & ", Customer: " & flxDC.TextMatrix(i, 2) & ", Check#:" & flxDC.TextMatrix(i, 3) & ", Bank:" & flxDC.TextMatrix(i, 4) & ", Date:" & flxDC.TextMatrix(i, 5) & ", Amt: " & flxDC.TextMatrix(i, 6)
                    flxRemarks.AddItem IIf(i = 0, "Deposited Checks: ", Space$(5)) & "PR#: " & flxDC.TextMatrix(i, 0) & ", PR Date:" & flxDC.TextMatrix(i, 1) & ", Customer: " & flxDC.TextMatrix(i, 2) & ", Check#:" & flxDC.TextMatrix(i, 3) & ", Bank:" & flxDC.TextMatrix(i, 4) & ", Date:" & flxDC.TextMatrix(i, 5) & ", Amt: " & flxDC.TextMatrix(i, 6)
                Next i
            End If
        End If
        'retrieve OR issued to PDC checks
        If flxOR.Rows > 0 Then
            If (flxRC.Rows + flxDC.Rows + flxOR.Rows) >= 5 Then
                myRem = IIf(myRem <> "", myRem & vbCrLf, "")
                myRem = myRem & "OR issued to Post-dated Checks: " & flxOR.Rows & " O.R. issued to PDC (Please see attachment for details)"
                Attch = True
            Else
                myRem = IIf(myRem <> "", myRem & vbCrLf & "OR issued to Post-dated Checks:", "OR issued to Post-dated Checks:")
                For i = 0 To flxOR.Rows - 1
                    myRem = myRem & vbCrLf & Space$(5) & "OR#:" & flxOR.TextMatrix(i, 0) & ", OR Amt:" & flxOR.TextMatrix(i, 1) & ", PR#:" & flxOR.TextMatrix(i, 2) & ", PR Date:" & flxOR.TextMatrix(i, 3) & ", Customer: " & flxOR.TextMatrix(i, 4) & ", Check#:" & flxOR.TextMatrix(i, 5) & ", Bank:" & flxOR.TextMatrix(i, 6) & ", Date:" & flxOR.TextMatrix(i, 7)
                    flxRemarks.AddItem IIf(i = 0, "OR issued to PDC: ", Space$(5)) & "OR#:" & flxOR.TextMatrix(i, 0) & ", OR Amt:" & flxOR.TextMatrix(i, 1) & ", PR#:" & flxOR.TextMatrix(i, 2) & ", PR Date:" & flxOR.TextMatrix(i, 3) & ", Customer: " & flxOR.TextMatrix(i, 4) & ", Check#:" & flxOR.TextMatrix(i, 5) & ", Bank:" & flxOR.TextMatrix(i, 6) & ", Date:" & flxOR.TextMatrix(i, 7)
                Next i
            End If
        End If
        If txtRemarks.Text <> "" Then flxRemarks.AddItem txtRemarks.Text 'myRem = myRem & vbCrLf & txtRemarks.Text
        .AddItem myRem & vbTab & "str"
        .AddItem txtPrepBy.Text & vbTab & "str"
        .AddItem txtAprvBy.Text & vbTab & "str"
        .AddItem Format$(CFdate.Value, "Dddd, Mmmm dd, yyyy") & vbTab & "str"
        .AddItem "Date Printed: " & Format$(Now, "mm/dd/yyyy hh:mm AMPM") & vbTab & "str"
        .AddItem "[SR#:" & SalesReturnSeries & vbTab & "str"
        .AddItem "[OR#:" & GetOR(PartsOR) & "]" & vbTab & "str"
        .AddItem GetFieldValue(SPonAcct)
        .AddItem "[OR#:" & GetOR(DpOR) & "]" & vbTab & "str"
        .AddItem GetFieldValue(DPamt)
        .AddItem "Add CV Check " & txtCVchk.Text & IIf(txtCVno.Text <> "", "(CV NO. " & txtCVno.Text & ")", "") & vbTab & "str"
        .AddItem Val(Replace(txtCVamt.Text, ",", ""))
        If flxRemarks.Rows >= 2 Then
            .AddItem flxRemarks.TextMatrix(1, 0)
        Else
            .AddItem ""
        End If
        If flxRemarks.Rows >= 3 Then
            .AddItem flxRemarks.TextMatrix(2, 0)
        Else
            .AddItem ""
        End If
        If flxRemarks.Rows >= 4 Then
            .AddItem flxRemarks.TextMatrix(3, 0)
        Else
            .AddItem ""
        End If
        If flxRemarks.Rows >= 5 Then
            .AddItem flxRemarks.TextMatrix(4, 0)
        Else
            .AddItem ""
        End If
        If flxRemarks.Rows >= 6 Then
            .AddItem flxRemarks.TextMatrix(5, 0)
        Else
            .AddItem ""
        End If
        
    End With
    
    With crCashier
        .EnableParameterPrompting = False
        For i = 1 To flxPV.Rows - 1
            .ParameterFields(i).ClearCurrentValueAndRange
            If i = 45 Then
                If flxRemarks.Rows >= 1 Then
                    .ParameterFields(i).AddCurrentValue flxRemarks.TextMatrix(0, 0)
                Else
                    .ParameterFields(i).AddCurrentValue ""
                End If
            ElseIf i = 57 Then
                If flxRemarks.Rows >= 2 Then
                    .ParameterFields(i).AddCurrentValue flxRemarks.TextMatrix(1, 0)
                Else
                    .ParameterFields(i).AddCurrentValue ""
                End If
            ElseIf i = 58 Then
                If flxRemarks.Rows >= 3 Then
                    .ParameterFields(i).AddCurrentValue flxRemarks.TextMatrix(2, 0)
                Else
                    .ParameterFields(i).AddCurrentValue ""
                End If
            ElseIf i = 59 Then
                If flxRemarks.Rows >= 4 Then
                    .ParameterFields(i).AddCurrentValue flxRemarks.TextMatrix(3, 0)
                Else
                    .ParameterFields(i).AddCurrentValue ""
                End If
            ElseIf i = 60 Then
                If flxRemarks.Rows >= 5 Then
                    .ParameterFields(i).AddCurrentValue flxRemarks.TextMatrix(4, 0)
                Else
                    .ParameterFields(i).AddCurrentValue ""
                End If
            ElseIf i = 61 Then
                If flxRemarks.Rows >= 6 Then
                    .ParameterFields(i).AddCurrentValue flxRemarks.TextMatrix(5, 0)
                Else
                    .ParameterFields(i).AddCurrentValue ""
                End If
            Else
                .ParameterFields(i).AddCurrentValue IIf(flxPV.TextMatrix(i, 1) <> "str", Val(flxPV.TextMatrix(i, 0)), Left$(flxPV.TextMatrix(i, 0), 255))
            End If
        Next i
        CRViewer1.ReportSource = crCashier
        CRViewer1.ViewReport
        CRViewer1.Zoom 75
        Call CRViewer1.PrintReport
    End With
    If flxChk.Rows > 6 Or (flxRC.Rows + flxDC.Rows + flxOR.Rows) >= 5 Then
        'AQry = "SELECT CFATTACHMENT.* From CFATTACHMENT WHERE (((CFATTACHMENT.CFDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
        'mySQL
        AQry = "SELECT CFATTACHMENT.* From CFATTACHMENT WHERE CFATTACHMENT.CFDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
        
        If CtrlNumExist(AQry) Then
            'Display Report
            SSTab1.TabEnabled(3) = True
            Call PrintPreviewAttachment
        Else
            Call UpdateAttachment
        '    'Display Report
            SSTab1.TabEnabled(3) = True
            Call PrintPreviewAttachment
        End If
    End If
    SSTab1.Tab = 2
End Sub

Private Sub UpdateAttachment()
    Dim i As Integer
    
    If flxChk.Rows > 6 Then
        For i = 1 To flxChk.Rows - 1
            'LYRcon.Execute "INSERT INTO CFATTACHMENT ( ATYPE, CHECKNO, CUSTOMER, BANK, CHKDATE, CHKAMT, CFDATE )" & _
                "VALUES ('CURRENT CHECKS', '" & flxChk.TextMatrix(i, 4) & "', '" & flxChk.TextMatrix(i, 1) & "', '" & flxChk.TextMatrix(i, 3) & "', #" & flxChk.TextMatrix(i, 2) & "#, " & flxChk.TextMatrix(i, 6) & ", #" & Format$(CFdate.Value, "MM/DD/YYYY") & "#);"
            'mySQL
            LYRcon.Execute "INSERT INTO CFATTACHMENT ( ATYPE, CHECKNO, CUSTOMER, BANK, CHKDATE, CHKAMT, CFDATE )" & _
                "VALUES ('CURRENT CHECKS', '" & flxChk.TextMatrix(i, 4) & "', '" & flxChk.TextMatrix(i, 1) & "', '" & flxChk.TextMatrix(i, 3) & "', '" & Format(flxChk.TextMatrix(i, 2), "yyyy-mm-dd") & "', " & flxChk.TextMatrix(i, 6) & ", '" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
            
        Next i
    End If
    If (flxRC.Rows + flxDC.Rows + flxOR.Rows) >= 5 Then
        If flxRC.Rows >= 1 Then
            For i = 0 To flxRC.Rows - 1
                'LYRcon.Execute "INSERT INTO CFATTACHMENT ( ATYPE, CHECKNO, CUSTOMER, BANK, CHKDATE, CHKAMT, PRNO, PRDATE, CFDATE )" & _
                    "VALUES ('RECEIVED CHECKS', '" & flxRC.TextMatrix(i, 3) & "', '" & flxRC.TextMatrix(i, 2) & "', '" & flxRC.TextMatrix(i, 4) & "', #" & flxRC.TextMatrix(i, 5) & "#, " & PureNum(flxRC.TextMatrix(i, 6)) & ", '" & flxRC.TextMatrix(i, 0) & "', #" & flxRC.TextMatrix(i, 1) & "#, #" & CFdate.Value & "#);"
                'mySQL
                LYRcon.Execute "INSERT INTO CFATTACHMENT ( ATYPE, CHECKNO, CUSTOMER, BANK, CHKDATE, CHKAMT, PRNO, PRDATE, CFDATE )" & _
                    "VALUES ('RECEIVED CHECKS', '" & flxRC.TextMatrix(i, 3) & "', '" & flxRC.TextMatrix(i, 2) & "', '" & flxRC.TextMatrix(i, 4) & "', '" & Format(flxRC.TextMatrix(i, 5), "yyyy-mm-dd") & "', " & Val(Replace(flxRC.TextMatrix(i, 6), ",", "")) & ", '" & flxRC.TextMatrix(i, 0) & "', '" & Format(flxRC.TextMatrix(i, 1), "yyyy-mm-dd") & "', '" & Format(CFdate.Value, "yyyy-mm-dd") & "';"
                
            Next i
        End If
        If flxDC.Rows >= 1 Then
            For i = 0 To flxDC.Rows - 1
                'LYRcon.Execute "INSERT INTO CFATTACHMENT ( ATYPE, CHECKNO, CUSTOMER, BANK, CHKDATE, CHKAMT, PRNO, PRDATE, CFDATE )" & _
                    "VALUES ('DEPOSITED CHECKS', '" & flxDC.TextMatrix(i, 3) & "', '" & flxDC.TextMatrix(i, 2) & "', '" & flxDC.TextMatrix(i, 4) & "', #" & flxDC.TextMatrix(i, 5) & "#, " & PureNum(flxDC.TextMatrix(i, 6)) & ", '" & flxDC.TextMatrix(i, 0) & "', #" & flxDC.TextMatrix(i, 1) & "#, #" & CFdate.Value & "#);"
                'mySQL
                LYRcon.Execute "INSERT INTO CFATTACHMENT ( ATYPE, CHECKNO, CUSTOMER, BANK, CHKDATE, CHKAMT, PRNO, PRDATE, CFDATE )" & _
                    "VALUES ('DEPOSITED CHECKS', '" & flxDC.TextMatrix(i, 3) & "', '" & flxDC.TextMatrix(i, 2) & "', '" & flxDC.TextMatrix(i, 4) & "', '" & Format(flxDC.TextMatrix(i, 5), "yyyy-mm-dd") & "', " & Val(Replace(flxDC.TextMatrix(i, 6), ",", "")) & ", '" & flxDC.TextMatrix(i, 0) & "', '" & Format(flxDC.TextMatrix(i, 1), "yyyy-mm-dd") & "', '" & Format(CFdate.Value, "yyyy-mm-dd") & "';"
                
            Next i
        End If
        If flxOR.Rows >= 1 Then
            For i = 0 To flxOR.Rows - 1
                'LYRcon.Execute "INSERT INTO CFATTACHMENT ( ATYPE, CHECKNO, CUSTOMER, BANK, CHKDATE, CHKAMT, PRNO, PRDATE, CFDATE )" & _
                    "VALUES ('O.R. ISSUED TO POST-DATED CHECKS', '" & flxOR.TextMatrix(i, 5) & "', '" & flxOR.TextMatrix(i, 4) & "', '" & flxOR.TextMatrix(i, 6) & "', #" & flxOR.TextMatrix(i, 7) & "#, " & PureNum(flxOR.TextMatrix(i, 1)) & ", '" & flxOR.TextMatrix(i, 2) & "', #" & flxOR.TextMatrix(i, 3) & "#, #" & CFdate.Value & "#);"
                'mySQL
                LYRcon.Execute "INSERT INTO CFATTACHMENT ( ATYPE, CHECKNO, CUSTOMER, BANK, CHKDATE, CHKAMT, PRNO, PRDATE, CFDATE )" & _
                    "VALUES ('O.R. ISSUED TO POST-DATED CHECKS', '" & flxOR.TextMatrix(i, 5) & "', '" & flxOR.TextMatrix(i, 4) & "', '" & flxOR.TextMatrix(i, 6) & "', '" & Format(flxOR.TextMatrix(i, 7), "yyyy-mm-dd") & "', " & Val(Replace(flxOR.TextMatrix(i, 1), ",", "")) & ", '" & flxOR.TextMatrix(i, 2) & "', '" & Format(flxOR.TextMatrix(i, 3), "yyyy-mm-dd") & "', '" & Format(CFdate.Value, "yyyy-mm-dd") & "';"
                
            Next i
        End If
    End If
End Sub

Private Sub PrintPreviewAttachment()
    Dim crAttach As New CFattachment
    
    'Qry = "SELECT CFATTACHMENT.* From CFattachment WHERE (((CFATTACHMENT.CFDATE)=#" & Format$(CFdate.Value, "MM/DD/YYYY") & "#));"
    'mySQL
    Qry = "SELECT CFATTACHMENT.* From CFattachment WHERE CFATTACHMENT.CFDATE='" & Format(Format$(CFdate.Value, "yyyy-mm-dd")) & "';"
    
    
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    crAttach.DiscardSavedData
    crAttach.Database.SetDataSource rsRPT
    crAttach.EnableParameterPrompting = False
    crAttach.ParameterFields(1).ClearCurrentValueAndRange
    crAttach.ParameterFields(1).AddCurrentValue txtPrepBy.Text
    crAttach.ParameterFields(2).ClearCurrentValueAndRange
    crAttach.ParameterFields(2).AddCurrentValue txtAprvBy.Text
    crAttach.ParameterFields(3).ClearCurrentValueAndRange
    crAttach.ParameterFields(3).AddCurrentValue SysBranch
    crAttach.ParameterFields(4).ClearCurrentValueAndRange
    crAttach.ParameterFields(4).AddCurrentValue SysAddress
    CRViewer2.ReportSource = crAttach
    CRViewer2.ViewReport
        
End Sub

Function DisplayIBP() As String
    Dim rsIBP As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    
    'Qry = "SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.REMARKS, PAYMENTS.ORNO, PAYMENTS.INSTALLMENT, PAYMENTS.DISCOUNT " & _
          "From PAYMENTS WHERE (((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((PAYMENTS.REMARKS) Like 'Paid%'));"
    'mySQL
    Qry = "SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.REMARKS, PAYMENTS.ORNO, PAYMENTS.INSTALLMENT, PAYMENTS.DISCOUNT " & _
          "From PAYMENTS WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND (PAYMENTS.REMARKS) Like 'Paid%';"
    
    With rsIBP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Tmp = ""
            Do While Not .EOF
                If Tmp = "" Then
                    Tmp = "IBP: Customer-" & !CustomerName & " " & !Remarks & " OR#" & !ORNO & ", Amount: " & Format$(!Installment, "###,##0.00") & IIf(!Discount > 0, " Rebate:" & !Discount, "")
                Else
                    Tmp = Tmp & vbCrLf & "IBP: Customer-" & !CustomerName & " " & !Remarks & " OR#" & !ORNO & ", Amount:" & Format$(!Installment, "###,##0.00") & IIf(!Discount > 0, " Rebate:" & !Discount, "")
                End If
                .MoveNext
            Loop
            DisplayIBP = Tmp
        Else
            DisplayIBP = ""
        End If
        .Close
    End With

End Function

Function CountDepositedChk() As Integer
    Dim rsChk As New ADODB.Recordset
    
    'Qry = "SELECT Count(CHECKS.PRNO) AS CountOfPRNO From CHECKS " & _
          "WHERE (((CHECKS.DATE4DPSIT)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    'mySQL
    Qry = "SELECT Count(CHECKS.PRNO) AS CountOfPRNO From CHECKS " & _
          "WHERE CHECKS.DATE4DPSIT='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
    
    With rsChk
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CountDepositedChk = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            CountDepositedChk = 0
        End If
        .Close
    End With

End Function

Function CountReceiveChk() As Integer
    Dim rsChk As New ADODB.Recordset
    
    Qry = "SELECT Count(CHECKS.PRNO) AS CountOfPRNO From CHECKS " & _
          "WHERE CHECKS.PRDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"

    With rsChk
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CountReceiveChk = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            CountReceiveChk = 0
        End If
        .Close
    End With
End Function

Function CountIBP() As Integer
    Dim rsIBP As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT Count(PAYMENTS.ORNO) AS CountOfORNO FROM EMPLOYEES INNER JOIN PAYMENTS ON EMPLOYEES.EMPID = PAYMENTS.EMPID " & _
          "WHERE (((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((EMPLOYEES.POSITION)='Cashier') AND ((PAYMENTS.REMARKS) Like 'Paid%'));"
    'mySQL
    Qry = "SELECT Count(PAYMENTS.ORNO) AS CountOfORNO FROM EMPLOYEES INNER JOIN PAYMENTS ON EMPLOYEES.EMPID = PAYMENTS.EMPID " & _
          "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND EMPLOYEES.POSITION='Cashier' AND (PAYMENTS.REMARKS) Like 'Paid%';"
    
    With rsIBP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CountIBP = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            CountIBP = 0
        End If
        .Close
    End With
End Function

Private Sub DispORChkinFlx()
    Dim rsCheck As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    Dim sRow As Integer, LRow As Integer, i As Integer
    
'old    Qry = "SELECT CHECKALLOC.[OR NUMBER], CHECKALLOC.[OR AMOUNT], CHECKS.PRNO, CHECKS.PRDATE, CHECKS.CHKACCTNAME AS CUSTOMER, CHECKS.CHECKNO, CHECKS.BANK, CHECKS.CHECKDATE " & _
          "FROM CHECKALLOC INNER JOIN CHECKS ON CHECKALLOC.CHKSPRID = CHECKS.CHKSPRID WHERE (((CHECKALLOC.[OR DATE])=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
'old    Qry = "SELECT CHECKALLOC.[ORNO], CHECKALLOC.[ORAMT], CHECKS.PRNO, CHECKS.PRDATE, CUSTOMER.CUSTOMERNAME, CHECKS.CHECKNO, CHECKS.BANK, CHECKS.CHECKDATE " & _
          "FROM (CHECKALLOC INNER JOIN CHECKS ON CHECKALLOC.CHKSPRID = CHECKS.CHKSPRID) INNER JOIN CUSTOMER ON CHECKS.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((CHECKALLOC.[ORDATE])=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    
    'Qry = "SELECT CHECKALLOC.ORNO, CHECKALLOC.ORAMT, CHECKS.PRNO, CHECKS.PRDATE, CUSTOMER.CUSTOMERNAME, CHECKS.CHECKNO, CHECKS.BANK, CHECKS.CHECKDATE " & _
          "FROM (CHECKS INNER JOIN CHECKALLOC ON CHECKS.CHKSPRID = CHECKALLOC.CHKSPRID) INNER JOIN CUSTOMER ON CHECKALLOC.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
          "WHERE (((CHECKALLOC.ORDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((CHECKS.STATUS)<>'CURRENT-CLEARED'));"
    'mySQL
    Qry = "SELECT CHECKALLOC.ORNO, CHECKALLOC.ORAMT, CHECKS.PRNO, CHECKS.PRDATE, CUSTOMER.CUSTOMERNAME, CHECKS.CHECKNO, CHECKS.BANK, CHECKS.CHECKDATE " & _
          "FROM (CHECKS INNER JOIN CHECKALLOC ON CHECKS.CHKSPRID = CHECKALLOC.CHKSPRID) INNER JOIN CUSTOMER ON CHECKALLOC.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
          "WHERE CHECKALLOC.ORDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND CHECKS.STATUS<>'CURRENT-CLEARED';"

    With rsCheck
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxOR.Rows = 0
            flxCF.AddItem "O.R. issued to PDC:"
            sRow = flxCF.Rows - 1
            LRow = sRow
            .MoveFirst
            Do While Not .EOF
                Tmp = "OR#:" & .Fields("ORNO") & " OR Amt:" & Format$(.Fields("ORAmt"), "###,##0.00") & " PR#:" & !PRno & " PR Date:" & Format$(!PRDate, "mm/dd/yy") & " CUSTOMER:" & !CustomerName & " CHECK#:" & !CheckNo & "/" & !Bank & "/" & Format$(!CheckDate, "mm/dd/yy")
                flxCF.AddItem Tmp & vbTab & Tmp & vbTab & Tmp & vbTab & Tmp
                flxOR.AddItem .Fields("ORNO") & vbTab & Format$(.Fields("ORAmt"), "###,##0.00") & vbTab & !PRno & vbTab & Format$(!PRDate, "mm/dd/yyyy") & vbTab & !CustomerName & vbTab & !CheckNo & vbTab & !Bank & vbTab & Format$(!CheckDate, "mm/dd/yyyy")
                LRow = LRow + 1
                .MoveNext
            Loop
            flxCF.MergeCells = flexMergeFree
            For i = sRow To LRow
                flxCF.MergeRow(i) = True
            Next i
        End If
        .Close
    End With
End Sub

Private Sub DispDepChkinFlx()
    Dim rsCheck As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    Dim sRow As Integer, LRow As Integer, i As Integer
    
    'Qry = "SELECT CHECKS.PRNO, CHECKS.PRDATE, CHECKS.CUSTOMER, CHECKS.CHECKNO, CHECKS.BANK, CHECKS.CHECKDATE, CHECKS.AMOUNT " & _
          "From CHECKS WHERE (((CHECKS.DATE4DPSIT)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    'mySQL
    Qry = "SELECT CHECKS.PRNO, CHECKS.PRDATE, CHECKS.CUSTOMER, CHECKS.CHECKNO, CHECKS.BANK, CHECKS.CHECKDATE, CHECKS.AMOUNT " & _
          "From CHECKS WHERE CHECKS.DATE4DPSIT='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
    
    With rsCheck
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxDC.Rows = 0
            flxCF.AddItem "Deposited Checks:"
            sRow = flxCF.Rows - 1
            LRow = sRow
            .MoveFirst
            Do While Not .EOF
                Tmp = "PR#: " & !PRno & " CUSTOMER: " & !CUSTOMER & " CHECK#: " & !CheckNo & "/" & !Bank & "/" & Format$(!CheckDate, "Mmm dd, yyyy") & "/" & Format$(!Amount, "###,##0.00")
                flxCF.AddItem Tmp & vbTab & Tmp & vbTab & Tmp & vbTab & Tmp
                flxDC.AddItem !PRno & vbTab & Format$(!PRDate, "mm/dd/yyyy") & vbTab & !CUSTOMER & vbTab & !CheckNo & vbTab & !Bank & vbTab & Format$(!CheckDate, "mm-dd-yyyy") & vbTab & Format$(!Amount, "###,##0.00")
                LRow = LRow + 1
                .MoveNext
            Loop
            flxCF.MergeCells = flexMergeFree
            For i = sRow To LRow
                flxCF.MergeRow(i) = True
            Next i
        End If
        .Close
    End With
End Sub


Private Sub DispRecChkinFlx()
    Dim rsCheck As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    Dim sRow As Integer, LRow As Integer, i As Integer
    
    'Qry = "SELECT CHECKS.PRNO, CHECKS.PRDATE, CUSTOMER.CUSTOMERNAME, CHECKS.CHECKNO, CHECKS.BANK, CHECKS.CHECKDATE, CHECKS.AMOUNT " & _
          "FROM CHECKS INNER JOIN CUSTOMER ON CHECKS.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((CHECKS.PRDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((CHECKS.DATE4BRKDWN) Is Null) AND ((CHECKS.STATUS)<>'CANCELLED'));"
    'mySQL
    Qry = "SELECT CHECKS.PRNO, CHECKS.PRDATE, CUSTOMER.CUSTOMERNAME, CHECKS.CHECKNO, CHECKS.BANK, CHECKS.CHECKDATE, CHECKS.AMOUNT " & _
          "FROM CHECKS INNER JOIN CUSTOMER ON CHECKS.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE CHECKS.PRDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND ((CHECKS.DATE4BRKDWN) Is Null) AND CHECKS.STATUS<>'CANCELLED';"
    
    With rsCheck
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxRC.Rows = 0
            flxCF.AddItem "Received Checks:"
            sRow = flxCF.Rows - 1
            LRow = sRow
            .MoveFirst
            Do While Not .EOF
                Tmp = "PR#: " & !PRno & " CUSTOMER: " & !CustomerName & " CHECK#: " & !CheckNo & "/" & !Bank & "/" & Format$(!CheckDate, "Mmm dd, yyyy") & "/" & Format$(!Amount, "###,##0.00")
                flxCF.AddItem Tmp & vbTab & Tmp & vbTab & Tmp & vbTab & Tmp
                flxRC.AddItem !PRno & vbTab & CFdate.Value & vbTab & !CustomerName & vbTab & !CheckNo & vbTab & !Bank & vbTab & Format$(!CheckDate, "mm-dd-yyyy") & vbTab & Format$(!Amount, "###,##0.00")
                LRow = LRow + 1
                .MoveNext
            Loop
            flxCF.MergeCells = flexMergeFree
            For i = sRow To LRow
                flxCF.MergeRow(i) = True
            Next i
        End If
        .Close
    End With
End Sub

Private Sub DisplayIBPinFlx()
    Dim rsIBP As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    Dim sRow As Integer, LRow As Integer, i As Integer
    
    'Qry = "SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.REMARKS, PAYMENTS.ORNO, PAYMENTS.INSTALLMENT, PAYMENTS.DISCOUNT " & _
          "From PAYMENTS WHERE (((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((PAYMENTS.REMARKS) Like 'Paid%'));"
    'mySQL
    Qry = "SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.REMARKS, PAYMENTS.ORNO, PAYMENTS.INSTALLMENT, PAYMENTS.DISCOUNT " & _
          "From PAYMENTS WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND (PAYMENTS.REMARKS) Like 'Paid%';"
    
    With rsIBP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            sRow = flxCF.Rows - 1
            LRow = sRow
            .MoveFirst
            Do While Not .EOF
                Tmp = "IBP: Customer-" & !CustomerName & " " & !Remarks & " OR#" & !ORNO & ", Amount:" & Format$(!Installment, "###,##0.00") & IIf(!Discount > 0, " Rebate:" & !Discount, "")
                flxCF.AddItem Tmp & vbTab & Tmp & vbTab & Tmp
                LRow = LRow + 1
                .MoveNext
            Loop
            flxCF.MergeCells = flexMergeFree
            For i = sRow To LRow
                flxCF.MergeRow(i) = True
            Next i
        End If
        .Close
    End With
End Sub

Function AMTforPDC() As Single
    Dim rsChkPDC As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    
    'Qry = "SELECT Sum(CHECKALLOC.[ORAMT]) AS SumOfORAMT FROM CHECKS INNER JOIN CHECKALLOC ON CHECKS.CHKSPRID = CHECKALLOC.CHKSPRID " & _
          "WHERE (((CHECKALLOC.ORDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((CHECKS.DATE4BRKDWN) Is Null));"
    'mySQL
    Qry = "SELECT Sum(CHECKALLOC.ORAMT) AS SumOfORAMT FROM CHECKS INNER JOIN CHECKALLOC ON CHECKS.CHKSPRID = CHECKALLOC.CHKSPRID " & _
          "WHERE CHECKALLOC.ORDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND (CHECKS.DATE4BRKDWN) Is Null;"
    
    With rsChkPDC
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            AMTforPDC = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            AMTforPDC = 0
        End If
        .Close
    End With
End Function

Function ORforPDC() As String
    Dim rsChkPDC As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    
    'Qry = "SELECT CHECKALLOC.ORNO, CHECKS.DATE4BRKDWN FROM CHECKALLOC INNER JOIN CHECKS ON CHECKALLOC.CHKSPRID = CHECKS.CHKSPRID " & _
          "WHERE (((CHECKALLOC.ORDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((CHECKS.DATE4BRKDWN) Is Null));"
    'mySQL
    Qry = "SELECT CHECKALLOC.ORNO, CHECKS.DATE4BRKDWN FROM CHECKALLOC INNER JOIN CHECKS ON CHECKALLOC.CHKSPRID = CHECKS.CHKSPRID " & _
          "WHERE CHECKALLOC.ORDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND (CHECKS.DATE4BRKDWN) Is Null;"
    
    With rsChkPDC
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Tmp = ""
            Do While Not .EOF
                Tmp = IIf(Tmp = "", "[" & .Fields(0), Tmp & ", " & .Fields(0))
                .MoveNext
            Loop
            Tmp = Tmp & "]"
            ORforPDC = Tmp
        Else
            ORforPDC = ""
        End If
        .Close
    End With
End Function

Function getManager() As String
    Dim rsMgr As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS MANAGER, EMPLOYEES.POSITION " & _
          "From EMPLOYEES WHERE (((EMPLOYEES.POSITION)='BRANCH MANAGER'));"
    'mySQL
    Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS MANAGER, EMPLOYEES.POSITION " & _
          "From EMPLOYEES WHERE EMPLOYEES.POSITION='BRANCH MANAGER';"
    
    With rsMgr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            getManager = .Fields(0)
        Else
            getManager = ""
        End If
        .Close
    End With
End Function

Function CVNOS(myRem As String) As String
    Dim rsCV As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    
    'Qry = "SELECT CASHVOUCHER.VOUCHERNO, CASHVOUCHER.CVDATE, CASHVOUCHER.REMARKS From CASHVOUCHER " & _
          "WHERE (((CASHVOUCHER.CVDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((CASHVOUCHER.REMARKS)='" & myRem & "'));"
    'mySQL
    Qry = "SELECT CASHVOUCHER.VOUCHERNO, CASHVOUCHER.CVDATE, CASHVOUCHER.REMARKS From CASHVOUCHER " & _
          "WHERE CASHVOUCHER.CVDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND CASHVOUCHER.REMARKS='" & myRem & "';"
    
    With rsCV
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Tmp = ""
            Do While Not .EOF
                Tmp = IIf(Tmp = "", Format(.Fields(0), "0000"), Tmp & ", " & Format(.Fields(0), "0000"))
                .MoveNext
            Loop
        Else
            Tmp = ""
        End If
        .Close
    End With
    CVNOS = Tmp
End Function

Function PromoAmt() As Single
    Dim rsAmt As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT Sum(PROMO.CVAMOUNT) AS SumOfCVAMOUNT From PROMO " & _
          "WHERE (((PROMO.CVDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    'mySQL
    Qry = "SELECT Sum(PROMO.CVAMOUNT) AS SumOfCVAMOUNT From PROMO " & _
          "WHERE PROMO.CVDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
    
    With rsAmt
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            PromoAmt = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            PromoAmt = 0
        End If
        .Close
    End With

End Function

Function PromoPCV() As String       'Not applicable
    Dim rsPromo As New ADODB.Recordset
    Dim Qry As String, tmpNo As String, tmpDesc As String
    
    'Qry = "SELECT PROMO.PROMODESC, PROMO.CVNUMBER " & _
          "From PROMO WHERE (((PROMO.CVDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    'mySQL
    Qry = "SELECT PROMO.PROMODESC, PROMO.CVNUMBER " & _
          "From PROMO WHERE PROMO.CVDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
    
    With rsPromo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                tmpNo = IIf(tmpNo = "", !PromoDesc & " PCV#" & !CVNUMBER, tmpNo & ", " & !CVNUMBER)
                .MoveNext
            Loop
            PromoPCV = tmpNo
        Else
            PromoPCV = ""
        End If
        .Close
    End With

End Function
'mySQL
Function GetCVamt(myRem As String) As Single
    Dim rsCVamt As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT CASHVOUCHER.VOUCHERNO From CASHVOUCHER " & _
          "WHERE (((CASHVOUCHER.CVDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((CASHVOUCHER.REMARKS)='" & myRem & "'));"
    'mySQL
    Qry = "SELECT CASHVOUCHER.VOUCHERNO From CASHVOUCHER " & _
          "WHERE CASHVOUCHER.CVDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND CASHVOUCHER.REMARKS='" & myRem & "';"
    
    With rsCVamt
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetCVamt = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            GetCVamt = 0
        End If
        .Close
    End With
End Function

Function ExpenseTotal() As Single
    Dim rsCVsum As New ADODB.Recordset
    Dim cvQry As String
    
    'cvQry = "SELECT Sum(CASHVOUCHER.AMOUNT) AS SumOfAMOUNT " & _
            "From CASHVOUCHER WHERE (((CASHVOUCHER.CVDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((CASHVOUCHER.REMARKS)='EXPENSE'));"
    
    'mySQL
    cvQry = "SELECT Sum(CASHVOUCHER.AMOUNT) AS SumOfAMOUNT " & _
            "From CASHVOUCHER WHERE CASHVOUCHER.CVDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND CASHVOUCHER.REMARKS='EXPENSE';"
    With rsCVsum
        .Open cvQry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ExpenseTotal = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            ExpenseTotal = 0
        End If
        .Close
    End With
End Function


Function FlxTotal(X1 As Integer, X2 As Integer, y As Integer) As Currency
    Dim i As Integer, Total As Currency
    
    Total = 0
    For i = X1 To X2
        Total = Total + PureNum(flxCF.TextMatrix(i, y))
    Next i
    FlxTotal = Total
End Function

Function SrvPartsCollection(myField As String) As Single
    Dim rsService As New ADODB.Recordset
    Dim Qry As String
    Dim Ttl As Single
    
    Qry = "SELECT Sum(PAYMENTS." & myField & ") AS SumOfSERVICEPARTS " & _
          "From PAYMENTS WHERE (((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    
    'mySQL
    Qry = "SELECT Sum(PAYMENTS." & myField & ") AS SumOfSERVICEPARTS " & _
          "From PAYMENTS WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
    
    With rsService
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            SrvPartsCollection = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            SrvPartsCollection = 0
        End If
        .Close
    End With
End Function

Function CollectorCollection() As Single
    Dim rsColl As New ADODB.Recordset
    Dim Qry As String
    Dim x As Integer, Ttl As Single
    
    'Qry = "SELECT Sum(PAYMENTS.CASH) AS SumOfCASH, Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT, Sum(PAYMENTS.INSTALLMENT) AS SumOfINSTALLMENT, Sum(PAYMENTS.OTHERS) AS SumOfOTHERS " & _
          "FROM PAYMENTS INNER JOIN EMPLOYEES ON PAYMENTS.EMPID = EMPLOYEES.EMPID " & _
          "WHERE (((PAYMENTS.PAYDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((EMPLOYEES.POSITION)<>'Cashier'));"
    'mySQL
    Qry = "SELECT Sum(PAYMENTS.CASH) AS SumOfCASH, Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT, Sum(PAYMENTS.INSTALLMENT) AS SumOfINSTALLMENT, Sum(PAYMENTS.OTHERS) AS SumOfOTHERS " & _
          "FROM PAYMENTS INNER JOIN EMPLOYEES ON PAYMENTS.EMPID = EMPLOYEES.EMPID " & _
          "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND EMPLOYEES.POSITION<>'Cashier';"
    
    With rsColl
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Ttl = 0
            For x = 0 To 3
                Ttl = Ttl + IIf(IsNull(.Fields(x)), 0, .Fields(x))
            Next x
        End If
        .Close
    End With
    CollectorCollection = Ttl

End Function

Function GetOR(Qry As String) As String
    Dim rsCashOR As New ADODB.Recordset
    Dim MyBig As String, ORseries As String, LastSer As String
    Dim x As Integer
    
    With rsCashOR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ORseries = ""
            x = 1
            Do While Not .EOF
               If x = 1 Then
                    MyBig = .Fields(0)
                    ORseries = .Fields(0)
                Else
                    If Val(.Fields(0)) = Val(MyBig) + 1 Then
                        MyBig = Val(.Fields(0))
                        If Right$(ORseries, 1) <> "-" Then
                            ORseries = ORseries & "-"
                            LastSer = .Fields(0)
                        End If
                    Else
                        If Right$(ORseries, 1) = "-" Then
                            ORseries = ORseries & IIf(Left$(LastSer, 3) <> Left$(Format$(MyBig, "00000"), 3), Format(MyBig, "00000"), Right$(MyBig, 2)) & ", " & .Fields(0)
                            MyBig = .Fields(0)
                        Else
                            MyBig = .Fields(0)
                            ORseries = ORseries & ", " & .Fields(0)
                        End If
                    End If
                End If
                x = x + 1
                .MoveNext
            Loop
            If Right$(ORseries, 1) = "-" Then ORseries = ORseries & MyBig
        End If
        .Close
    End With
    GetOR = ORseries
End Function

Function CashierCollection() As Currency
    Dim rsCashColl As New ADODB.Recordset
    Dim Qry As String
    Dim x As Integer
    Dim Ttl As Double
    
    Qry = "SELECT Sum(PAYMENTS.CASH) AS SumOfCASH, Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT, Sum(PAYMENTS.INSTALLMENT) AS SumOfINSTALLMENT, Format((Sum([PAYMENTS].[OTHERS])),'Fixed') AS SumOfOTHERS " & _
          "FROM EMPLOYEES INNER JOIN PAYMENTS ON EMPLOYEES.EMPID = PAYMENTS.EMPID " & _
          "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND EMPLOYEES.POSITION='Cashier' AND (PAYMENTS.REMARKS) Not Like 'Paid%'; "
    With rsCashColl
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Ttl = 0
            For x = 0 To 3
                Ttl = Ttl + IIf(IsNull(.Fields(x)), 0, .Fields(x))
            Next x
        End If
        .Close
    End With
    CashierCollection = Ttl
End Function

Private Sub DisplayCFChecks()
    Dim rsBDchecks As New ADODB.Recordset
    Dim Qry As String
    
    'old Qry = "SELECT CHECKS.CHKACCTNAME, CHECKS.CHECKDATE, CHECKS.BANK, CHECKS.CHECKNO, CHECKS.PAIDTO, CHECKS.AMOUNT, CHECKS.DATE4BRKDWN, CHECKS.CHKSPRID " & _
          "From CHECKS WHERE (((CHECKS.DATE4BRKDWN)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    'old Qry = "SELECT CHECKS.CHKACCTNAME, CHECKS.CHECKDATE, CHECKS.BANK, CHECKS.CHECKNO, CHECKS.PAIDTO, CHECKS.AMOUNT, CHECKS.DATE4BRKDWN, CHECKS.CHKSPRID " & _
          "From CHECKS WHERE CHECKS.DATE4BRKDWN='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
    'mySQL
    Qry = "SELECT CHECKS.CHKACCTNAME, CHECKS.CHECKDATE, CHECKS.BANK, CHECKS.CHECKNO, CHECKS.PAIDTO, CHECKS.AMOUNT, CHECKS.DATE4BRKDWN, CHECKS.CHKSPRID " & _
          "From CHECKS WHERE (((CHECKS.DATE4BRKDWN)='" & Format$(CFdate.Value, "yyyy-mm-dd") & "') AND ((CHECKS.STATUS)<>'CANCELLED' Or (CHECKS.STATUS) Is Null));"

    With rsBDchecks
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxChk.Rows = 1
            Do While Not .EOF
                flxChk.AddItem vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & .Fields(3) & vbTab & .Fields(4) & vbTab & Format$(.Fields(5), "##,###,##0.00") & vbTab & !CHKSPRID
                .MoveNext
            Loop
            ChkLstRow = flxChk.Rows - 1
        Else
            flxChk.Rows = 1
            flxChk.Rows = 2
        End If
        .Close
    End With
End Sub

Private Sub DisplayCFDetails()
    Dim rsCF As New ADODB.Recordset
    Dim Qry As String, i As Integer
    Dim MyBills() As String, MyCoins() As String
    
    'Qry = "SELECT CASHFLOWMGR.* From CashFlowMgr " & _
          "WHERE (((CASHFLOWMGR.CFDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    'mySQL
    Qry = "SELECT CASHFLOWMGR.* From CashFlowMgr " & _
          "WHERE CASHFLOWMGR.CFDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
    With rsCF
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyBills = Split(.Fields("Bills"), "*")
            For i = 0 To 6
                txtBill(i).Text = MyBills(i)
            Next i
            MyCoins = Split(.Fields("Coins"), "*")
            For i = 0 To 5
                txtCoin(i).Text = MyCoins(i)
            Next i
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtCVchk.Text = IIf(IsNull(!PCVCHECK), "", !PCVCHECK)
            txtPrepBy.Text = .Fields("PreparedBy")
            txtAprvBy.Text = .Fields("ApprovedBy")
        End If
        .Close
    End With
End Sub

Public Sub ViewCashFlow()
    Dim i As Integer
    
    If NewCF Then
        If CashFlowExist Then
            MsgBox "Cash Flow for " & Format$(CFdate.Value, "Mmm. dd, yyyy") & " already exist!" & vbCrLf & _
                   "Please select a new date to process.", vbExclamation + vbOKOnly, "Warning"
            Frame1.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
        Else
            Call LockData(False)
        End If
    Else
        If CashFlowExist Then
            flxCF.Rows = 0
            Call DisplayCFChecks
            Call DisplayCFDetails
            Call ProcessCashFlow
            txtPrepBy.Locked = True
            txtAprvBy.Locked = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 1
        Else
            For i = 0 To 5
                txtBill(i).Text = ""
                txtCoin(i).Text = ""
            Next i
            txtBill(6).Text = ""
            txtPrepBy.Text = ""
            txtAprvBy.Text = ""
            flxCF.Rows = 0
            flxPV.Rows = 0
            flxChk.Rows = 1
            flxChk.Rows = 2
            SSTab1.Tab = 0
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            MsgBox "No Cash Flow made for " & Format$(CFdate, "Mmmm dd, yyyy"), vbExclamation + vbOKOnly, "Cash Flow"
            
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("This will update attachments that was previously not existing in Report!" & vbCrLf & vbCrLf & _
        "Do you want to continue?", vbQuestion + vbYesNo, "Update Attachment")
    If Ans = vbNo Then Exit Sub
    'LYRcon.Execute "DELETE CFATTACHMENT.*, From CFattachment WHERE (((CFATTACHMENT.CFDATE)=#" & CFdate.Value & "#));"
    'mySQL
    LYRcon.Execute "DELETE From CFattachment WHERE CFATTACHMENT.CFDATE='" & Format(CFdate.Value, "yyyy-mm-dd") & ";"
    
    Call UpdateAttachment
    Call PrintPreviewAttachment
        
End Sub
Private Sub cmdRefresh_Click()
    Call ProcessCashFlow
End Sub
Private Sub cmdTrans_Click(Index As Integer)
    Dim rsCF As New ADODB.Recordset
    Dim Ans As Integer
    Dim myMsg As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'AddNew
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=62));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=62;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            Call ClearEntries
            txtPrepBy.Text = GetSettings("OIC")
            txtAprvBy.Text = GetSettings("MANAGER")
            'txtNotedBy.Text = GetSettings("OIC")
            SETdate.MySource = "MGR CASHFLOW NEW"
            SETdate.Show vbModal
            If CashFlowExist Then
                MsgBox "Manager's Cash Flow for " & Format$(CFdate.Value, "Mmm. dd, yyyy") & " already exist!" & vbCrLf & _
                       "You may Edit the Cash Flow using the Edit Button", vbExclamation + vbOKOnly, "Warning"
                Frame1.Enabled = False
                Frame2.Enabled = False
                Frame3.Enabled = False
                Exit Sub
            ElseIf Not CashierCFExist Then
                MsgBox "Cashier's Cash Flow for " & Format$(CFdate.Value, "Mmm. dd, yyyy") & " does not exist!" & vbCrLf & _
                       "You have to build Cashier's Cash Flow before you can continue.", vbExclamation + vbOKOnly, "Warning"
                Frame1.Enabled = False
                Frame2.Enabled = False
                Frame3.Enabled = False
                Exit Sub
            Else
                Call LockData(False)
            End If
            Call tbSaveCancel(True)
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 0
            NewCF = True
            'Load Entry from Cashier's Cashflow if present
            'If Not CashierCFExist Then Exit Sub
            Call DisplayCFChecks
            Call DisplayCFDetails
            Call ProcessCashFlow
            txtPrepBy.Locked = True
            txtAprvBy.Locked = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
        
        Case 1      'Edit
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=62;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            If txtMgrCFId.Text = "" Then
                MsgBox "No Manager's Cash Flow available for Edit!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            NewCF = False
            Call LockData(False)
            Call tbSaveCancel(True)
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 0
        
        Case 2      'View
            Frame4.Enabled = True
            NewCF = False
            SETdate.MySource = "MGR CASHFLOW NEW"
            SETdate.Show vbModal
            If Not CashFlowExist Then
                MsgBox "Manager's Cash Flow for " & Format$(CFdate.Value, "Mmm. dd, yyyy") & " does not exist!", vbExclamation + vbOKOnly, "Warning"
                Exit Sub
            End If
            
            'Qry = "SELECT CASHFLOWMGR.MGRCFID From CashFlowMgr WHERE (((CASHFLOWMGR.CFDATE)=#" & Format$(CFdate.Value, "Mmm. dd, yyyy") & "#));"
            'mySQL
            Qry = "SELECT CASHFLOWMGR.MGRCFID From CashFlowMgr WHERE CASHFLOWMGR.CFDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
            rsCF.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsCF.RecordCount > 0 Then
                rsCF.MoveFirst
                txtMgrCFId.Text = rsCF!MgrCFid
            End If
            rsCF.Close
            'Load Entry from Cashier's Cashflow if present
            Call DisplayCFChecks
            Call DisplayCFDetails
            Call ProcessCashFlow
            txtPrepBy.Locked = True
            txtAprvBy.Locked = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
        Case 3      'print preview
            If txtTotalBills.Text = "" Then
                MsgBox "No Cash Flow available for preview!", vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Call ProcessCF
            SSTab1.TabEnabled(2) = True
        Case 4      'quit
            Unload Me
            MotoRep.tbMain.Visible = True
        
        Case 5      'save
            myMsg = ValidateCF
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Cash Flow?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveCashFlow
            Call tbSaveCancel(False)
            Call LockData(True)
            NewCF = False
            'Load Entry from Cashier's Cashflow if present
            Call DisplayCFChecks
            Call DisplayCFDetails
            Call ProcessCashFlow
            txtPrepBy.Locked = True
            txtAprvBy.Locked = True
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 1
            MsgBox "Cash Flow successfully saved!", vbInformation + vbOKOnly, "Save"
        Case 6      'cancel
            Ans = MsgBox("Do you want to cancel entry?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            Call LockData(True)
            Call tbSaveCancel(False)
            txtChk.Visible = False
            Call ClearEntries
    End Select

End Sub

Private Sub flxChk_Click()
    ActRow = flxChk.Row
    ActCol = flxChk.Col
    
    If ActCol = 0 Or ActRow <= ChkLstRow Then Exit Sub
    With txtChk
        .Top = flxChk.CellTop + flxChk.Top
        .Left = flxChk.CellLeft + flxChk.Left
        .Width = flxChk.CellWidth
        .Text = flxChk.Text
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
        Select Case ActCol
            Case 1, 3, 4
                .MaxLength = 30
            Case 5
                .MaxLength = 50
            Case 2, 6
                .MaxLength = 15
        End Select
    End With
End Sub

Private Sub flxChk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxChk_Click
End Sub

Private Sub flxChk_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ActCol <> 0 Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuChecks)
    End If
End Sub

Private Sub flxChk_Scroll()
    txtChk.Visible = False
End Sub

Private Sub Form_Load()
    Call LockData(True)
    Call tbSaveCancel(False)
    CFdate.Value = Now
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    MotoRep.sBar.Panels(6).Text = "MANAGER'S CASH FLOW"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsRPT.State = adStateOpen Then rsRPT.Close
End Sub

Private Sub mnuAdd_Click()
    With flxChk
        If .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 3) = "" Or .TextMatrix(.Rows - 1, 4) = "" Or .TextMatrix(.Rows - 1, 6) = "" Then
            MsgBox "Incomplete entry on last row's Check information.", vbExclamation + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        .AddItem ""
    End With
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    If flxChk.TextMatrix(ActRow, 7) <> "" Then
        MsgBox "Deletion not allowed for Check#" & flxChk.TextMatrix(ActRow, 4), vbExclamation + vbOKOnly, "Delete Cancelled"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to delete row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbNo Then Exit Sub
    If flxChk.Rows = 2 Then
        flxChk.Rows = 1
        flxChk.Rows = 2
    Else
        flxChk.RemoveItem ActRow
    End If
End Sub

Function CashFlowExist() As Boolean
    Dim rsCF As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT CASHFLOWMGR.* From CashFlowMgr " & _
          "WHERE (((CASHFLOWMGR.CFDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"

    'mySQL
    Qry = "SELECT CASHFLOWMGR.* From CashFlowMgr " & _
          "WHERE CASHFLOWMGR.CFDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"

    With rsCF
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CashFlowExist = True
        Else
            CashFlowExist = False
        End If
        .Close
    End With
End Function

Function CashierCFExist() As Boolean
    Dim rsCF As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT CashFlow.* From CashFlow " & _
          "WHERE (((CashFlow.CFDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#));"
    'mySQL
    Qry = "SELECT CashFlow.* From CashFlow " & _
          "WHERE CashFlow.CFDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "';"
    
    With rsCF
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CashierCFExist = True
        Else
            CashierCFExist = False
        End If
        .Close
    End With
End Function

Function ValidateCF() As String
    Dim MyErr As String
    
    MyErr = ""
    'If CashFlowExist Then MyErr = "Cash Flow already exist for " & Format$(CFdate, "Mmmm dd, yyyy")
    If PureNum(txtTotalBills.Text) = 0 Then MyErr = MyErr & vbCrLf & "Incomplete entry on Bills break down."
    If txtCVno.Text <> "" Then
        If txtCVchk.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Field CV Check Details"
    End If
    If txtPrepBy.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Field Prepared By."
    If txtAprvBy.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Field Approved By."
    With flxChk
        If .TextMatrix(1, 1) <> "" Or .TextMatrix(1, 3) <> "" Or .TextMatrix(1, 4) <> "" Or .TextMatrix(1, 6) <> "" Then
            If .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 3) = "" Or .TextMatrix(.Rows - 1, 4) = "" Or .TextMatrix(.Rows - 1, 6) = "" Then
                MyErr = MyErr & vbCrLf & "Incomplete entry on last row's Check information."
            End If
        End If
    End With
    ValidateCF = MyErr
End Function
'mySQL
Private Sub SaveCashFlow()
    Dim rsCashFlow As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim i As Integer, Cs As Integer, SaveCheck As Boolean
    Dim tmpBill As String, tmpCoin As String, Qry As String
    Dim MyID As Long
    
    'With rsCashFlow
    '    If NewCF Then
    '        .Open "CashFlowMgr", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
    '        Qry = "SELECT CASHFLOWMGR.* From CashFlowMgr WHERE (((CASHFLOWMGR.MGRCFID)=" & txtMgrCFId.Text & "));"
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
    '    End If
    '    tmpBill = ""
    '    For i = 0 To 6
    '        If tmpBill = "" Then
    '            tmpBill = IIf(txtBill(i).Text = "", "0", txtBill(i).Text)
    '        Else
    '            tmpBill = tmpBill & "*" & IIf(txtBill(i).Text = "", "0", txtBill(i).Text)
    '        End If
    '    Next i
    '    .Fields("Bills") = tmpBill
    '    tmpCoin = ""
    '    For i = 0 To 5
    '        If tmpCoin = "" Then
    '            tmpCoin = IIf(txtCoin(i).Text = "", "0", txtCoin(i).Text)
    '        Else
    '            tmpCoin = tmpCoin & "*" & IIf(txtCoin(i).Text = "", "0", txtCoin(i).Text)
    '        End If
    '    Next i
    '    !EXPENSECV = txtCVno.Text
    '    !EXPENSEAMT = Val(Replace(txtCVamt.Text, ",", ""))
    '    !CASHOUTCV = txtCashOutCv.Text
    '    !CASHOUTAMT = Val(txtCashOutAmt.Text)
    '    !LYRCV = txtLyrCv.Text
    '    !LYRCVAMT = Val(txtLyrAmt.Text)
    '    .Fields("Coins") = tmpCoin
    '    .Fields("CFDate") = Format$(CFdate.Value, "mm/dd/yyyy")
    '    .Fields("Remarks") = txtRemarks.Text
    '    .Fields("PreparedBy") = txtPrepBy.Text
    '    .Fields("ApprovedBy") = txtAprvBy.Text
    '    .Fields("UpdatedBy") = UserID
    '    .Fields("LastUpdate") = Date
    '    !PCVCHECK = txtCVchk.Text
    '    .Update
    '    MyId = !MgrCfId
    '    If NewCF Then txtMgrCFId.Text = MyId
    'End With
    'save checks [Algo:] All checks are loaded in flxChk with corresponding CheckID
    'Changes made to Checks with ID must be updated thru a variable that indicates changes
    'User can add Checks but can never remove checks with ID
    
    
    'mySQL
    tmpBill = ""
    For i = 0 To 6
        If tmpBill = "" Then
            tmpBill = IIf(txtBill(i).Text = "", "0", txtBill(i).Text)
        Else
            tmpBill = tmpBill & "*" & IIf(txtBill(i).Text = "", "0", txtBill(i).Text)
        End If
    Next i
    tmpCoin = ""
    For i = 0 To 5
        If tmpCoin = "" Then
            tmpCoin = IIf(txtCoin(i).Text = "", "0", txtCoin(i).Text)
        Else
            tmpCoin = tmpCoin & "*" & IIf(txtCoin(i).Text = "", "0", txtCoin(i).Text)
        End If
    Next i
     'mySQL here
    UpdtTag = ""
    UpdtTag = Format(Now, "yyyy-mm-dd hh:mm:ss")
    If NewCF Then
        LYRcon.Execute "Insert Into cashflowmgr( BILLS, EXPENSECV, EXPENSEAMT, CASHOUTCV, CASHOUTAMT, " & _
            "LYRCV, LYRCVAMT, COINS, CFDATE, REMARKS, PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE, " & _
            "PCVCHECK)VALUES( '" & tmpBill & "', '" & txtCVno.Text & "', " & Val(Replace(txtCVamt.Text, ",", "")) & ", " & _
            "'" & txtCashOutCv.Text & "', " & Val(txtCashOutAmt.Text) & ", '" & txtLyrCv.Text & "', " & _
            "" & Val(txtLyrAmt.Text) & ", '" & tmpCoin & "', '" & Format$(CFdate.Value, "yyyy-mm-dd") & "', " & _
            "'" & txtRemarks.Text & "', '" & txtPrepBy.Text & "', '" & txtAprvBy.Text & "', " & UserID & ", " & _
            "'" & UpdtTag & "', '" & txtCVchk.Text & "');"
        MyID = GetFieldValue("SELECT CASHFLOWMGR.MGRCFID From CashFlowMgr WHERE CASHFLOWMGR.LASTUPDATE='" & UpdtTag & "';")
        txtMgrCFId.Text = MyID
    Else
        LYRcon.Execute "Update CashFlowMgr SET BILLS = '" & tmpBill & "', EXPENSECV = '" & txtCVno.Text & "', " & _
            "EXPENSEAMT = " & Val(Replace(txtCVamt.Text, ",", "")) & ", CASHOUTCV = '" & txtCashOutCv.Text & "', " & _
            "CASHOUTAMT = '" & txtCashOutCv.Text & "', LYRCV = '" & Val(txtCashOutAmt.Text) & "', " & _
            "LYRCVAMT = " & Val(txtLyrCv.Text) & ", COINS = '" & tmpCoin & "', CFDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "', " & _
            "REMARKS = '" & txtRemarks.Text & "', PREPAREDBY = '" & txtPrepBy.Text & "', APPROVEDBY = '" & txtAprvBy.Text & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & UpdtTag & "', PCVCHECK = '" & txtCVchk.Text & "' " & _
            "Where CashFlowMgr.LASTUPDATE = '" & UpdtTag & "';"
    End If
    
    'With flxChk
    '    For i = 1 To .Rows - 1
    '        SaveCheck = False
    '        If .TextMatrix(i, 7) <> "" And .TextMatrix(i, 8) = "CHANGE" Then
    '            Qry = "SELECT CHECKS.* From CHECKS WHERE (((CHECKS.CHKSPRID)=" & .TextMatrix(i, 7) & "));"
    '            rsCheck.Open "Checks", LYRcon, adOpenKeyset, adLockOptimistic
    '            SaveCheck = True
    '        ElseIf .TextMatrix(i, 7) = "" And .TextMatrix(i, 1) <> "" And .TextMatrix(i, 4) <> "" And .TextMatrix(i, 6) <> "" Then
    '            rsCheck.Open "Checks", LYRcon, adOpenKeyset, adLockOptimistic
    '            rsCheck.AddNew
    '            SaveCheck = True
    '        End If
    '        If SaveCheck Then
    '            rsCheck!ChkAcctName = .TextMatrix(i, 1)
    '            rsCheck!CheckDate = .TextMatrix(i, 2)
    '            rsCheck!Bank = .TextMatrix(i, 3)
    '            rsCheck!CheckNo = .TextMatrix(i, 4)
    '            rsCheck!PAIDTO = .TextMatrix(i, 5)
    '            rsCheck!Amount = PureNum(.TextMatrix(i, 6))
    '            rsCheck!DATE4BRKDWN = Format$(CFdate.Value, "mm/dd/yyyy")
    '            rsCheck!MgrCfId = MyId
    '            rsCheck!UPDATEDBY = UserID
    '            rsCheck!LASTUPDATE = Date
    '            rsCheck.Update
    '            .TextMatrix(i, 7) = rsCheck!CHKSPRID
    '            rsCheck.Close
    '        End If
    '    Next i
    'End With
    
    'mySQL here
    With flxChk
        For i = 1 To .Rows - 1
            SaveCheck = False
            UpdtTag = ""
            UpdtTag = Format(Now, "yyyy-mm-dd hh:mm:ss")
            If .TextMatrix(i, 7) <> "" And .TextMatrix(i, 8) = "CHANGE" Then
                SaveCheck = True
                LYRcon.Execute "Update checks SET CHKACCTNAME = '" & .TextMatrix(i, 1) & "', CHECKDATE = '" & Format(.TextMatrix(i, 2), "yyyy-mm-dd") & "', " & _
                    "BANK = '" & .TextMatrix(i, 3) & "', CHECKNO = '" & .TextMatrix(i, 4) & "', PAIDTO = '" & .TextMatrix(i, 5) & "', AMOUNT = " & Val(Replace(.TextMatrix(i, 6), ",", "")) & ", " & _
                    "DATE4BRKDWN = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "', MGRCFID = " & MyID & ", UPDATEDBY = " & UserID & ", LASTUPDATE = '" & UpdtTag & "' " & _
                    "where checks.CHKSPRID='" & .TextMatrix(i, 7) & "';"
            ElseIf .TextMatrix(i, 7) = "" And .TextMatrix(i, 1) <> "" And .TextMatrix(i, 4) <> "" And .TextMatrix(i, 6) <> "" Then
                SaveCheck = True
                LYRcon.Execute "Insert Into checks( CHKACCTNAME, CHECKDATE, BANK, CHECKNO, PAIDTO, AMOUNT, DATE4BRKDWN, MGRCFID, UPDATEDBY, " & _
                    "LASTUPDATE)VALUES( '" & .TextMatrix(i, 1) & "', '" & Format(.TextMatrix(i, 2), "yyyy-mm-dd") & "', '" & .TextMatrix(i, 3) & "', " & _
                    "'" & .TextMatrix(i, 4) & "', '" & .TextMatrix(i, 5) & "', " & Val(Replace(.TextMatrix(i, 6), ",", "")) & ", '" & Format$(CFdate.Value, "yyyy-mm-dd") & "', " & _
                    "" & MyID & ", " & UserID & ", '" & UpdtTag & "');"
            End If
            If SaveCheck = True Then .TextMatrix(i, 7) = GetFieldValue("Select checks.CHKSPRID from CHECKS WHERE CHECKS.LASTUPDATE='" & UpdtTag & "';")
        Next i
    End With
End Sub

Function SalesReturnAmt() As Currency
    Dim rsRet As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    
    'Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE (((SRETURNSP.RETDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((SRETURNSP.PROCESS)='REPLACE') AND ((SPINVOICE.INVTYPE)='CASH'));"

    'mySQL
    Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE SRETURNSP.RETDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND SRETURNSP.PROCESS='REPLACE' AND SPINVOICE.INVTYPE='CASH';"

    With rsRet
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            SalesReturnAmt = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            SalesReturnAmt = 0
        End If
        .Close
    End With
    
End Function

Function SalesReturnSeries() As String
    Dim rsSalesRet As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    
    'Qry = "SELECT SRETURNSP.RETNO FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE (((SRETURNSP.RETDATE)=#" & Format$(CFdate.Value, "mm/dd/yyyy") & "#) AND ((SRETURNSP.PROCESS)='REPLACE') AND ((SPINVOICE.INVTYPE)='CASH'));"
    'mySQL
    Qry = "SELECT SRETURNSP.RETNO FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE SRETURNSP.RETDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND SRETURNSP.PROCESS='REPLACE' AND SPINVOICE.INVTYPE='CASH';"
    
    With rsSalesRet
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Tmp = ""
            Do While Not .EOF
                Tmp = IIf(Tmp = "", "[" & .Fields(0), Tmp & ", " & .Fields(0))
                .MoveNext
            Loop
            Tmp = Tmp & "]"
            SalesReturnSeries = Tmp
        Else
            SalesReturnSeries = ""
        End If
        .Close
    End With
    
End Function

Private Sub RXResizer1_ResizeComplete()
    Label3(4).Left = (Me.Width - 100) - Label3(4).Width
    Label3(5).Left = (Me.Width - 130) - Label3(5).Width
    With flxCF
        .ColWidth(0) = (.Width - 300) * 0.2
        .ColWidth(1) = (.Width - 300) * 0.5
        .ColWidth(2) = (.Width - 300) * 0.15
        .ColWidth(3) = (.Width - 300) * 0.15
    End With
    With flxChk
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.2
        .ColWidth(2) = (.Width - 300) * 0.14
        .ColWidth(3) = (.Width - 300) * 0.2
        .ColWidth(4) = (.Width - 300) * 0.14
        .ColWidth(5) = (.Width - 300) * 0.16
        .ColWidth(6) = (.Width - 300) * 0.14
        .ColWidth(7) = 0
        .ColWidth(8) = 0
    End With
End Sub


Private Sub txtBill_Change(Index As Integer)
    Dim Ans As Single
    
    Ans = Val(txtBill(Index).Text) * Val(lblBill(Index).Caption)
    lblBillTtl(Index).Caption = Format$(Ans, "##,###,##0.00")
    Call ComputeBills
End Sub

Private Sub txtBill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyDown Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtBill_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtChk_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = IIf(ActCol = 2, "0123456789/", "0123456789.")
    If KeyAscii > 26 And (ActCol = 2 Or ActCol = 6) Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxChk.SetFocus
        If ActCol = 6 Then
            flxChk.Col = 0
        Else
            flxChk.Col = ActCol + 1
        End If
    End If
End Sub

Private Sub txtChk_LostFocus()
    If ActCol = 2 And txtChk.Text <> "" And Not IsDate(txtChk.Text) Then
        MsgBox "Invalid entry on Check Date.", vbExclamation + vbOKOnly, "Validation Error"
        txtChk.SetFocus
        txtChk.SelStart = 0
        txtChk.SelLength = Len(txtChk.Text)
        Exit Sub
    End If
    flxChk.TextMatrix(ActRow, ActCol) = txtChk.Text
    txtChk.Visible = False
    If flxChk.TextMatrix(ActRow, 7) <> "" Then flxChk.TextMatrix(ActRow, 8) = "CHANGE"
    
End Sub

Private Sub txtCoin_Change(Index As Integer)
    Dim Ans As Single
    
    Ans = Val(txtCoin(Index).Text) * Val(lblCoin(Index).Caption)
    lblCoinTtl(Index).Caption = Format$(Ans, "##,###,##0.00")
    Call ComputeCoins
End Sub

Private Sub txtCoin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        SendKeys "+{TAB}"
    ElseIf KeyCode = vbKeyDown Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCoin_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


