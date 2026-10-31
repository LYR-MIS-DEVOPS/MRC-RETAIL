VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form CashFlowMgr 
   Caption         =   "Manager's Cash Flow"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   -36060
   ClientWidth     =   12630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   12630
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCashOutCv 
      Height          =   285
      Left            =   1680
      TabIndex        =   109
      Top             =   10800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCashOutAmt 
      Height          =   285
      Left            =   2280
      TabIndex        =   108
      Top             =   10800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCFid 
      Height          =   285
      Left            =   360
      TabIndex        =   107
      Top             =   10800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxCollector 
      Height          =   615
      Left            =   120
      TabIndex        =   84
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxRemarks 
      Height          =   735
      Left            =   2160
      TabIndex        =   80
      Top             =   9960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxOR 
      Height          =   735
      Left            =   10320
      TabIndex        =   76
      Top             =   9960
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
      Left            =   8520
      TabIndex        =   75
      Top             =   10080
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
      Left            =   6120
      TabIndex        =   74
      Top             =   9960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   0
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
      Caption         =   "Build Cash Flow"
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
      Image           =   "CashFlowMgrNew.frx":0000
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPV 
      Height          =   615
      Left            =   3840
      TabIndex        =   68
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   19
      Tag             =   "XYF"
      Top             =   720
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   14843
      _Version        =   393216
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
      TabPicture(0)   =   "CashFlowMgrNew.frx":059A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tabTOF"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Print Preview"
      TabPicture(1)   =   "CashFlowMgrNew.frx":05B6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CRViewer1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Attachment"
      TabPicture(2)   =   "CashFlowMgrNew.frx":05D2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CRViewer2"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab tabTOF 
         Height          =   4335
         Left            =   2520
         TabIndex        =   92
         Top             =   12260
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         ForeColor       =   255
         TabCaption(0)   =   "RECEIVED FUNDS"
         TabPicture(0)   =   "CashFlowMgrNew.frx":05EE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdCancelTOF"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdOkTOF"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "flxTOF"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtTtlTOF"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtCVdate"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtCVNum"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtRecFrom"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         Begin VB.TextBox txtRecFrom 
            Height          =   300
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   100
            Top             =   3480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtCVNum 
            Height          =   300
            Left            =   1800
            TabIndex        =   99
            Top             =   3480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtCVdate 
            Height          =   300
            Left            =   480
            TabIndex        =   98
            Top             =   3480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTtlTOF 
            Height          =   375
            Left            =   5520
            TabIndex        =   94
            Top             =   3240
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid flxTOF 
            Height          =   2775
            Left            =   120
            TabIndex        =   93
            Top             =   480
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   4895
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   315
            BackColorBkg    =   -2147483633
            FormatString    =   "    |^C.V. DATE       |^ C.V. NUMBER    |< RECEIVED FROM                        |> AMOUNT         "
         End
         Begin lvButton.lvButtons_H cmdOkTOF 
            Height          =   420
            Left            =   120
            TabIndex        =   96
            Top             =   3840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   741
            Caption         =   "OK"
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
            Image           =   "CashFlowMgrNew.frx":060A
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCancelTOF 
            Height          =   420
            Left            =   1575
            TabIndex        =   97
            Top             =   3840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   741
            Caption         =   "Cancel"
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
            Image           =   "CashFlowMgrNew.frx":0BA4
            cBack           =   -2147483633
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "TOTAL:"
            Height          =   255
            Left            =   4320
            TabIndex        =   95
            Top             =   3300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   240
         TabIndex        =   77
         Tag             =   "XYF"
         Top             =   6480
         Width           =   11895
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
            Left            =   5160
            MaxLength       =   50
            TabIndex        =   103
            Tag             =   "XYF"
            Top             =   1040
            Width           =   6615
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
            Left            =   5160
            MaxLength       =   20
            TabIndex        =   102
            Tag             =   "XYF"
            Top             =   240
            Width           =   1695
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
            Left            =   5160
            MaxLength       =   20
            TabIndex        =   101
            Tag             =   "XYF"
            Top             =   640
            Width           =   1695
         End
         Begin lvButton.lvButtons_H cmdTOF 
            Height          =   375
            Left            =   480
            TabIndex        =   91
            Tag             =   "xyf"
            Top             =   1010
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Rec. Funds"
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
            cBack           =   -2147483633
         End
         Begin VB.TextBox txtRecTOF 
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   90
            Tag             =   "XYF"
            Top             =   640
            Width           =   1335
         End
         Begin VB.TextBox txtUndpColl 
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
            Left            =   1920
            TabIndex        =   86
            Tag             =   "XYF"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtRemarks 
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
            Left            =   5160
            MaxLength       =   200
            TabIndex        =   82
            Tag             =   "XYF"
            Top             =   1440
            Width           =   6615
         End
         Begin VB.ComboBox cboCashier 
            Height          =   315
            Left            =   8400
            TabIndex        =   81
            Tag             =   "XYF"
            Top             =   240
            Width           =   3375
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
            Left            =   8400
            TabIndex        =   13
            Tag             =   "XYF"
            Top             =   640
            Width           =   3375
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
            Index           =   5
            Left            =   3480
            TabIndex        =   106
            Tag             =   "XYF"
            Top             =   1070
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
            Index           =   4
            Left            =   3960
            TabIndex        =   105
            Tag             =   "XYF"
            Top             =   270
            Width           =   1095
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
            Left            =   3960
            TabIndex        =   104
            Tag             =   "XYF"
            Top             =   670
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Received Funds:"
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
            TabIndex        =   89
            Tag             =   "XYF"
            Top             =   670
            Width           =   1695
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Undeposited Collection:"
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
            TabIndex        =   85
            Tag             =   "XYF"
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
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
            Left            =   3360
            TabIndex        =   83
            Tag             =   "XYF"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Prepared By:"
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
            Left            =   6840
            TabIndex        =   79
            Tag             =   "XYF"
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
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
            Left            =   6840
            TabIndex        =   78
            Tag             =   "XYF"
            Top             =   670
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   240
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   480
         Width           =   11895
         Begin MSComCtl2.DTPicker CFdate 
            Height          =   315
            Left            =   9120
            TabIndex        =   65
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
            Format          =   129957891
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
            Left            =   8160
            TabIndex        =   66
            Tag             =   "XYF"
            Top             =   210
            Width           =   1815
         End
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
         Height          =   2295
         Left            =   240
         TabIndex        =   55
         Tag             =   "XYF"
         Top             =   4200
         Width           =   11895
         Begin VB.TextBox txtChk 
            Height          =   285
            Left            =   120
            TabIndex        =   67
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid flxChk 
            Height          =   1695
            Left            =   120
            TabIndex        =   56
            Tag             =   "XY"
            Top             =   240
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   2990
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   315
            BackColorBkg    =   -2147483633
            ScrollBars      =   2
            BorderStyle     =   0
            FormatString    =   $"CashFlowMgrNew.frx":113E
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
         Height          =   2940
         Left            =   6360
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   1200
         Width           =   5775
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
            Left            =   3480
            TabIndex        =   64
            Tag             =   "XY"
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   7
            Tag             =   "XY"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   8
            Tag             =   "XY"
            Top             =   555
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   2
            Left            =   720
            TabIndex        =   9
            Tag             =   "XY"
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   3
            Left            =   720
            TabIndex        =   10
            Tag             =   "XY"
            Top             =   1185
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   4
            Left            =   720
            TabIndex        =   11
            Tag             =   "XY"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox txtCoin 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   5
            Left            =   720
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
            Left            =   3480
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
            Left            =   3480
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
            Left            =   3480
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
            Left            =   3480
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
            Left            =   3480
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
            Left            =   3480
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
            Left            =   2040
            TabIndex        =   54
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
            Left            =   2040
            TabIndex        =   53
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
            Left            =   2040
            TabIndex        =   52
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
            Left            =   2040
            TabIndex        =   51
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
            Left            =   2040
            TabIndex        =   50
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
            Left            =   2040
            TabIndex        =   49
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
            Left            =   2400
            TabIndex        =   48
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
            Left            =   2400
            TabIndex        =   47
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
            Left            =   2400
            TabIndex        =   46
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
            Left            =   2400
            TabIndex        =   45
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
            Left            =   2400
            TabIndex        =   44
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
            Left            =   2400
            TabIndex        =   43
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
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   1200
         Width           =   5775
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
            TabIndex        =   14
            Tag             =   "XY"
            Top             =   2475
            Width           =   1575
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   6
            Left            =   600
            TabIndex        =   6
            Tag             =   "XY"
            Top             =   2130
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   5
            Left            =   600
            TabIndex        =   5
            Tag             =   "XY"
            Top             =   1815
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   4
            Left            =   600
            TabIndex        =   4
            Tag             =   "XY"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   3
            Left            =   600
            TabIndex        =   3
            Tag             =   "XY"
            Top             =   1185
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   2
            Left            =   600
            TabIndex        =   2
            Tag             =   "XY"
            Top             =   870
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   1
            Left            =   600
            TabIndex        =   1
            Tag             =   "XY"
            Top             =   555
            Width           =   1215
         End
         Begin VB.TextBox txtBill 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   0
            Left            =   600
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            Left            =   2280
            TabIndex        =   34
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
            Left            =   2280
            TabIndex        =   33
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
            Left            =   2280
            TabIndex        =   32
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
            Left            =   2280
            TabIndex        =   31
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
            Left            =   2280
            TabIndex        =   30
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
            Left            =   2280
            TabIndex        =   29
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
            Left            =   2280
            TabIndex        =   28
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
            Left            =   1920
            TabIndex        =   27
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
            Left            =   1920
            TabIndex        =   26
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
            Left            =   1920
            TabIndex        =   25
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
            Left            =   1920
            TabIndex        =   24
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
            Left            =   1920
            TabIndex        =   23
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
            Left            =   1920
            TabIndex        =   22
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
            Left            =   1920
            TabIndex        =   21
            Tag             =   "XY"
            Top             =   240
            Width           =   375
         End
      End
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   87
         Tag             =   "xy"
         Top             =   480
         Width           =   12135
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
      Begin CRVIEWERLibCtl.CRViewer CRViewer2 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   88
         Tag             =   "xy"
         Top             =   480
         Width           =   12135
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
            Picture         =   "CashFlowMgrNew.frx":11D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowMgrNew.frx":176A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowMgrNew.frx":1D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowMgrNew.frx":229E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowMgrNew.frx":2838
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CashFlowMgrNew.frx":2DD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   1
      Left            =   1815
      TabIndex        =   70
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   741
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
      Image           =   "CashFlowMgrNew.frx":336C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   2
      Left            =   3630
      TabIndex        =   71
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
      Image           =   "CashFlowMgrNew.frx":3906
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   3
      Left            =   5445
      TabIndex        =   72
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
      Image           =   "CashFlowMgrNew.frx":3EA0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   4
      Left            =   7260
      TabIndex        =   73
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
      Image           =   "CashFlowMgrNew.frx":443A
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   1440
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12870
      m_MDI_Height    =   9720
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
      Left            =   9480
      TabIndex        =   18
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
      Left            =   9510
      TabIndex        =   17
      Top             =   510
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LYRGCOM"
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
      TabIndex        =   16
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LYRGCOM"
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
      TabIndex        =   15
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
   Begin VB.Menu mnuTOF 
      Caption         =   "Transfer of Funds"
      Visible         =   0   'False
      Begin VB.Menu mnuAddTOF 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuDelTOF 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "CashFlowMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_Proj As CRAXDRT.Application
Dim CF_Prvw As CRAXDRT.Report
Dim CF_Attc As CRAXDRT.Report

Dim mcLyr As CRAXDRT.Application
Dim mcReport As CRAXDRT.Report
Dim rsRPT As New ADODB.Recordset
Dim ActRow As Integer, ActCol As Integer
Dim ChkLstRow As Integer
Dim NewCF As Boolean, Brn As String
Dim SPcash As Currency, SPcashWS As Currency, SvcAmt As Currency, DPamt As Currency, OtherAmt As Currency, SPonAcct As Currency, SvcReg As Currency, SvcFSC As Currency, OthrAcct As Currency, TtlAR As Currency
Dim NetCash As Currency, ChkTtlAmt As Currency, BillsTtl As Currency, CoinsTtl As Currency, Cash4Deposit As Currency, CollectorAmt As Currency, TtlPDCamt As Currency
Dim CashReturn As Currency, TtlLessOR As Currency, CashOnHand As Currency, AdvLiq As Currency, DueToMfkg As Currency, Withdrawal As Currency, TOFgmo As Currency
Dim CashOutAmt As Currency, sRetAmt As Currency, OnLineAmt As Currency, CCardAmt As Currency, IBrnAmt As Currency, TtlLess As Currency, TtlColltn As Currency, NetColl As Currency
Dim OnLineSERVICEAmt As Currency, OnLineOthersAmt As Currency, OnLineSPAmt As Currency

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    Frame2.Enabled = Not bln
    Frame3.Enabled = Not bln
    'Frame4.Enabled = Not bln
    Frame5.Enabled = Not bln
End Sub

Private Sub ClearEntries()
    Dim i As Integer
    
    CFdate.Value = Now
    txtRemarks.Text = ""
    txtAprvby.Text = ""
    'txtNotedBy.Text = ""
    flxChk.Rows = 1
    flxChk.Rows = 2
    'flxCF.Rows = 0
    For i = 0 To 5
        txtBill(i).Text = ""
        txtCoin(i).Text = ""
    Next i
    txtBill(6).Text = ""
    txtTotalBills.Text = ""
    txtTotalCoins.Text = ""
    txtCVNO.Text = ""
    txtCVamt.Text = ""
    txtCVchk.Text = ""
End Sub

Private Sub tbSaveCancel(bln As Boolean)
    Dim x As Integer
    
    If bln Then
        For x = 0 To 2
            cmdTrans(x).Visible = False
        Next x
        cmdTrans(3).Left = 0
        cmdTrans(4).Left = cmdTrans(3).Width
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
    Else
        For x = 0 To 2
            cmdTrans(x).Visible = True
        Next x
        cmdTrans(4).Visible = False
        cmdTrans(3).Visible = False
    End If
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
                If IsNull(.Fields(0)) Then
                    .MoveNext
                Else
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
                End If
            Loop
            If Right$(ORseries, 1) = "-" Then ORseries = ORseries & MyBig
        End If
        .Close
    End With
    GetOR = ORseries
End Function


Private Sub DisplayCFChecks()
    Dim rsPdcDue As New ADODB.Recordset
    Dim rsCrnt As New ADODB.Recordset
    Dim rsEnCash As New ADODB.Recordset
    Dim rsChkRec As New ADODB.Recordset
    Dim Qry As String, i As Integer, SkipMe As Boolean
    
    flxChk.Rows = 1
    Qry = "SELECT checks.CHECKNO, checks.CHECKDATE, checks.CHKACCTNAME, checks.`STATUS`, checks.BANK, checks.AMOUNT, checks.CHKSPRID " & _
          "From checks Where checks.`STATUS` = 'CURRENT-CLEARED' AND checks.PRDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checks.PRNO = '';"
    With rsCrnt 'display current checks
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxChk.Rows = 1
            Do While Not .EOF
                '1-CHECK NO., 2-DATE,  3-CHECK ACCT. NAME, 4-CHECK STATUS,  5-BANK,  6-AMOUNT, 7 -checkid
                flxChk.AddItem vbTab & !CHECKNO & vbTab & Format(!CHECKDATE, "MM/DD/YYYY") & vbTab & !CHKACCTNAME & vbTab & "CURRENT" & vbTab & !BANK & vbTab & Format$(!Amount, "#,##0.00") & vbTab & !CHKSPRID
                .MoveNext
            Loop
            ChkLstRow = flxChk.Rows - 1
        Else
            flxChk.Rows = 1
            ChkLstRow = 0
        End If
        .Close
    End With
    'Identify PDC DUE
    Qry = "SELECT checks.CHECKNO, checks.CHKACCTNAME, checks.BANK, checks.AMOUNT, checks.`STATUS`, checks.CHKSPRID, checks.CHECKDATE From checks " & _
          "WHERE checks.`STATUS` = 'PDC DUE' AND checks.CHECKDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checks.PRNO <> '';"
    With rsPdcDue
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxChk.AddItem vbTab & !CHECKNO & vbTab & Format(!CHECKDATE, "MM/DD/YYYY") & vbTab & !CHKACCTNAME & vbTab & !STATUS & vbTab & !BANK & vbTab & Format$(!Amount, "#,##0.00") & vbTab & !CHKSPRID
                .MoveNext
            Loop
            ChkLstRow = flxChk.Rows - 1
        End If
        .Close
    End With
    
    'Display Checks for encashment
    Qry = "SELECT checks.CHECKNO, checks.CHECKDATE, checks.CHKACCTNAME, checks.`STATUS`, checks.BANK, checks.AMOUNT, checks.CHKSPRID From checks " & _
          "Where checks.`STATUS` LIKE '%ENCASHMENT%' AND checks.DATE4BRKDWN = '" & Format(CFdate.Value, "yyyy-mm-dd") & "';"
    With rsEnCash
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxChk.AddItem vbTab & !CHECKNO & vbTab & Format(!CHECKDATE, "MM/DD/YYYY") & vbTab & !CHKACCTNAME & vbTab & !STATUS & vbTab & !BANK & vbTab & Format$(!Amount, "#,##0.00") & vbTab & !CHKSPRID
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    'Display Received checks
    Qry = "SELECT checks.CHECKNO, checks.CHECKDATE, checks.CHKACCTNAME, checks.`STATUS`, checks.BANK, checks.AMOUNT, checks.CHKSPRID From checks " & _
          "Where checks.`STATUS` NOT IN ( 'CURRENT-CLEARED','CANCELLED') AND checks.DATE4BRKDWN = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checks.PRDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "';"
    flxRC.Rows = 0
    With rsChkRec
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxRC.AddItem vbTab & !CHECKNO & vbTab & Format(!CHECKDATE, "MM/DD/YYYY") & vbTab & !CHKACCTNAME & vbTab & !STATUS & vbTab & !BANK & vbTab & Format$(!Amount, "#,##0.00") & vbTab & !CHKSPRID
                .MoveNext
            Loop
        End If
        .Close
    End With
        

End Sub
Public Sub ShowRecord(MyCFid As Long)
    Dim rsCF As New ADODB.Recordset
    Dim rsFunds As New ADODB.Recordset
    Dim Qry As String, i As Integer, TtlFRec As Currency
    Dim MyBills() As String, MyCoins() As String
    
    Qry = "SELECT CASHFLOWMGR.* From CASHFLOWMGR " & _
          "WHERE CASHFLOWMGR.MGRCFID =" & MyCFid & ";"
    With rsCF
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CFdate.Value = !CFdate
            MyBills = Split(.Fields("Bills"), "*")
            For i = 0 To 6
                txtBill(i).Text = MyBills(i)
            Next i
            MyCoins = Split(.Fields("Coins"), "*")
            For i = 0 To 5
                txtCoin(i).Text = MyCoins(i)
            Next i
            cboCashier.Clear
            cboCashier.AddItem .Fields("PREPAREDBY")
            cboCashier.ItemData(cboCashier.ListCount - 1) = !UPDATEDBY
            cboCashier.ListIndex = 0
            txtAprvby.Text = .Fields("APPROVEDBY")
            'txtNotedBy.Text = IIf(IsNull(!NotedBy), "", !NotedBy)
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtCVchk.Text = IIf(IsNull(!PCVCHECK), "", !PCVCHECK)
            txtCVamt.Text = IIf(IsNull(!EXPENSEAMT), "", Format(!EXPENSEAMT, "#,##0.00"))
            txtCVNO.Text = IIf(IsNull(!EXPENSECV), "", !EXPENSECV)
            txtCFid.Text = !MGRCFID
        End If
        .Close
    End With
    'Retrieve Refunds Received
    
    Call DisplayCFChecks
    Call ShowReceivedFunds
    Call CFpreview
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    
End Sub

Private Sub CFdate_LostFocus()
    If NewCF Then
        Call DisplayCFChecks
        Call ShowReceivedFunds
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("This will update attachments that was previously not existing in Report!" & vbCrLf & vbCrLf & _
        "Do you want to continue?", vbQuestion + vbYesNo, "Update Attachment")
    If Ans = vbNo Then Exit Sub
    LYRcon.Execute "DELETE CFATTACHMENT.*, From CFattachment WHERE (((CFATTACHMENT.CFDATE)=#" & CFdate.Value & "#));"
    'Call UpdateAttachment
    'Call PrintPreviewAttachment
        
End Sub

Private Sub LoadCashier()
    Dim rsCashier As New ADODB.Recordset
    Dim MyID As Integer
        
    MyID = 0
    With rsCashier
        Qry = "SELECT concat(employees.LASTNAME, ', ', employees.FIRSTNAME, ' ',LEFT(employees.MIDDLENAME,1),'.') AS CASHIER, employees.EmpId " & _
              "From employees Where employees.`POSITION` = 'CASHIER' Order By  employees.LASTNAME;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboCashier.Clear
            Do While Not .EOF
                cboCashier.AddItem !Cashier
                cboCashier.ItemData(cboCashier.ListCount - 1) = !EMPID
                If UserID = !EMPID Then MyID = cboCashier.ListCount - 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    If cboCashier.ListCount > 0 Then cboCashier.ListIndex = MyID

End Sub

Private Sub cmdCancelTOF_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Cancel Received Funds?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    flxTOF.Rows = 1
    flxTOF.Rows = 2
    txtTtlTOF.Text = ""
    Call cmdOkTOF_Click
End Sub

Private Sub cmdOkTOF_Click()
    txtRecTOF.Text = txtTtlTOF.Text
    tabTOF.Visible = False
    tabTOF.Top = 12000
    Call LockData(False)
    cmdTrans(3).Enabled = True
    cmdTrans(4).Enabled = True
End Sub

Private Sub cmdTOF_Click()
    Call LockData(True)
    cmdTrans(3).Enabled = False
    cmdTrans(4).Enabled = False
    With tabTOF
        .Top = (SSTab1.Height - .Height) / 2
        .Left = (SSTab1.Width - .Width) / 2
        .Visible = True
    End With
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim rsFunds As New ADODB.Recordset
    Dim Ans As Integer
    Dim myMsg As String, MyAcc As String, Qry As String, myValidation As String
    
    Select Case Index
        Case 0      'Build
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=62;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If CashFlowExist Then
                MsgBox "Cash Flow for " & Format$(CFdate.Value, "Mmm. dd, yyyy") & " already exist!" & vbCrLf & _
                       "Please select a new date to process.", vbExclamation + vbOKOnly, "Warning"
                Frame1.Enabled = False
                Frame2.Enabled = False
                Frame3.Enabled = False
                Exit Sub
            Else
                Call LockData(False)
            End If
            
            Qry = "SELECT cashflow.CASHFLOWID From cashflow Where cashflow.CFDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "';"
            MyAcc = GetFieldValue(Qry)
            If MyAcc = "" Then
                MsgBox "NO available CASHIER'S CASHFLOW for " & Format(CFdate.Value, "MM/DD/YYYY") & "!" & vbCrLf & "Please select another Date to continue transaction...", vbExclamation + vbOKOnly, "Cashier's Cashflow Unavailable"
                Exit Sub
            End If
            
            Qry = "Select Group_concat(inv.TRANSNO) as TRANSNO From spinvoice as inv " & _
                  "Left Join (Select REFERENCE From spitr Where REMARKS IN ('CASH','CREDIT') and TRANSDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' Group by REFERENCE) as tr ON tr.REFERENCE = inv.TRANSNO " & _
                  "Where inv.INVDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' and inv.REMARKS NOT IN ('CANCELLED') and tr.REFERENCE is NULL;"
            If rsFunds.State = adStateOpen Then rsFunds.Close
            rsFunds.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsFunds.RecordCount > 0 Then
                rsFunds.MoveFirst
                Do While Not rsFunds.EOF
                    If IIf(IsNull(rsFunds!TransNo), "", rsFunds!TransNo) <> "" Then
                        myValidation = "This Invoice had no Spareparts Details Found: " & rsFunds!TransNo
                    End If
                    rsFunds.MoveNext
                Loop
                rsFunds.Close
            End If
            
            If myValidation <> "" Then
                MsgBox myValidation, vbInformation + vbOKOnly, "Validation"
                Exit Sub
            End If
            
            Call ClearEntries
            'txtPrepBy.Text = GetSettings("CASHIER")
            'txtAprvBy.Text = GetSettings("OIC")
            'txtNotedBy.Text = GetSettings("OIC")
            cboCashier.Locked = False
            txtAprvby.Locked = False
            Call tbSaveCancel(True)
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            'SSTab1.TabEnabled(3) = False
            SSTab1.Tab = 0
            NewCF = True
            Call DisplayCFChecks
            Call ShowReceivedFunds
            
        Case 1      'View
            FrmSummary.mySrc = "Manager's Cash Flow"
            FrmSummary.Show
            Unload Me
        Case 2      'quit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
        Case 3      'save
            myMsg = ValidateCF
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Cash Flow?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveCashFlow
            Call CFpreview   'ProcessCashFlow
            Call tbSaveCancel(False)
            SSTab1.Tab = 1
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = True
            MsgBox "Cash Flow successfully saved!", vbInformation + vbOKOnly, "Save"
        Case 4      'cancel
            Ans = MsgBox("Do you want to cancel entry?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            Call LockData(True)
            Call tbSaveCancel(False)
            txtChk.Visible = False
            Call ClearEntries
    End Select

End Sub

Private Sub ShowReceivedFunds()
    Dim rsFunds As New ADODB.Recordset
    Dim Qry As String, TtlRec As Currency
    
    'Retrieve Refunds Received
    Qry = "SELECT fundsreceived.CVDATE, fundsreceived.CVNO, fundsreceived.RECFROM, fundsreceived.AMOUNT " & _
          "From fundsreceived Where fundsreceived.CVDATE ='" & Format(CFdate.Value, "yyyy-mm-dd") & "';"
    With rsFunds
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxTOF.Rows = 1
            Do While Not .EOF
                flxTOF.AddItem vbTab & Format(!CVDate, "yyyy-mm-dd") & vbTab & !cvno & vbTab & !RecFrom & vbTab & Format(!Amount, "#,##0.00")
                TtlFRec = TtlFRec + !Amount
                .MoveNext
            Loop
        End If
        .Close
        txtRecTOF.Text = Format(TtlFRec, "#,##0.00")
    End With

End Sub


Private Sub flxChk_Click()
    ActRow = flxChk.Row
    ActCol = flxChk.Col
    
    If ActCol = 0 Or ActCol = 4 Or ActRow <= ChkLstRow Then Exit Sub
    With txtChk
        .Tag = flxChk.TextMatrix(ActRow, ActCol)
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
            Case 1, 3
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

Private Sub flxChk_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If ActCol <> 0 Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuChecks)
    End If
End Sub

Private Sub flxChk_Scroll()
    txtChk.Visible = False
End Sub

Private Sub flxTOF_Click()
    ActCol = flxTOF.Col
    ActRow = flxTOF.Row

    If ActRow = 0 Or ActCol = 0 Then Exit Sub
    Select Case ActCol
        Case 1  'CV Date
            With txtCVDATE
                .Tag = flxTOF.TextMatrix(ActRow, ActCol)
                .Top = flxTOF.CellTop + flxTOF.Top
                .Left = flxTOF.CellLeft + flxTOF.Left
                .Width = flxTOF.CellWidth
                .Text = flxTOF.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 2  'CV#
            With txtCVNum
                .Tag = flxTOF.TextMatrix(ActRow, ActCol)
                .Top = flxTOF.CellTop + flxTOF.Top
                .Left = flxTOF.CellLeft + flxTOF.Left
                .Width = flxTOF.CellWidth
                .Text = flxTOF.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 3  'Received From
            With txtRecFrom
                .Tag = flxTOF.TextMatrix(ActRow, ActCol)
                .Top = flxTOF.CellTop + flxTOF.Top
                .Left = flxTOF.CellLeft + flxTOF.Left
                .Width = flxTOF.CellWidth
                .Text = flxTOF.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 4  'CV Amount
            With txtCVNum
                .Tag = flxTOF.TextMatrix(ActRow, ActCol)
                .Top = flxTOF.CellTop + flxTOF.Top
                .Left = flxTOF.CellLeft + flxTOF.Left
                .Width = flxTOF.CellWidth
                .Text = flxTOF.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
    End Select

End Sub

Private Sub flxTOF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxTOF_Click
End Sub

Private Sub flxTOF_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If ActCol <> 0 Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuTOF)
    End If
End Sub

Private Sub flxTOF_Scroll()
    txtCVDATE.Visible = False
    txtCVNum.Visible = False
    txtRecFrom.Visible = False
End Sub

Private Sub Form_Load()
    Dim rsCashier As New ADODB.Recordset
    Dim MyID As Integer, Qry As String
    
    Call LockData(True)
    Call tbSaveCancel(False)
    CFdate.Value = Now
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    'SSTab1.TabEnabled(3) = False
    NewCF = False
    MyID = -1
    With rsCashier
        Qry = "SELECT concat(employees.FIRSTNAME,' ', IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')),' ', employees.LASTNAME) AS CASHIER, employees.EmpId " & _
              "From Employees Where employees.DEPARTMENT = 'ADMIN' Order By CASHIER"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboCashier.Clear
            Do While Not .EOF
                cboCashier.AddItem !Cashier
                cboCashier.ItemData(cboCashier.ListCount - 1) = !EMPID
                If UserID = !EMPID Then MyID = cboCashier.ListCount - 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    If cboCashier.ListCount > 0 And MyID >= 0 Then cboCashier.ListIndex = MyID
    'Qry = "SELECT transignatory.EMPLOYEE From transignatory Where transignatory.TRANSACTION = 'Manager Cashflow';"
    'txtAprvBy.Text = GetFieldValue(Qry)
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

Private Sub mnuAddTOF_Click()
    With flxTOF
        If .TextMatrix(.Rows - 1, 1) = "" Or .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 3) = "" Or .TextMatrix(.Rows - 1, 4) = "" Then
            MsgBox "Incomplete entry on last Row", vbExclamation + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        .AddItem ""
    End With
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    If ActRow <= ChkLstRow Then
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

Private Sub mnuDelTOF_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to delete row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbNo Then Exit Sub
    If flxTOF.Rows = 2 Then
        flxTOF.Rows = 1
        flxTOF.Rows = 2
    Else
        flxTOF.RemoveItem ActRow
    End If

End Sub


Function CashFlowExist() As Boolean
    Dim rsCF As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT CashFlowMgr.* From CashFlowMgr " & _
          "WHERE CashFlowMgr.CFDATE='" & Format$(CFdate.Value, "yyyy-MM-DD") & "';"
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

Function ValidateCF() As String
    Dim MyErr As String, MyAcc As String, Qry As String
    
    MyErr = ""
    Qry = "SELECT cashflow.CASHFLOWID From cashflow Where cashflow.CFDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "';"
    MyAcc = GetFieldValue(Qry)
    If MyAcc = "" Then MyErr = "Cashier's Cash Flow does not exist for " & Format$(CFdate, "Mmmm dd, yyyy")
    If CashFlowExist Then MyErr = "Cash Flow already exist for " & Format$(CFdate, "Mmmm dd, yyyy")
    If txtTotalBills.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Bills break down."
    If cboCashier.ListIndex = -1 Then MyErr = MyErr & vbCrLf & "Please select a Name from the list for PREPARED BY"
    If txtAprvby.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Field Approved By."
    With flxChk
        If .Rows > 1 Then
        If .TextMatrix(1, 1) <> "" Or .TextMatrix(1, 3) <> "" Or .TextMatrix(1, 4) <> "" Or .TextMatrix(1, 6) <> "" Then
            If .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 3) = "" Or .TextMatrix(.Rows - 1, 4) = "" Or .TextMatrix(.Rows - 1, 6) = "" Then
                MyErr = MyErr & vbCrLf & "Incomplete entry on last row's Check information."
            End If
        End If
        End If
    End With
    ValidateCF = MyErr
End Function

Private Sub SaveCashFlow()
    Dim rsCashFlow As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim i As Integer, Cs As Integer, SaveCheck As Boolean
    Dim tmpBill As String, tmpCoin As String
    Dim MyID As Long
    
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
        LYRcon.Execute "Insert Into cashflowmgr( BILLS, EXPENSECV, EXPENSEAMT, COINS, CFDATE, REMARKS, PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE, " & _
            "PCVCHECK)VALUES( '" & tmpBill & "', '" & txtCVNO.Text & "', " & Val(Replace(txtCVamt.Text, ",", "")) & ", " & _
            "'" & tmpCoin & "', '" & Format$(CFdate.Value, "yyyy-mm-dd") & "', " & _
            "'" & txtRemarks.Text & "', '" & cboCashier.Text & "', '" & txtAprvby.Text & "', " & UserID & ", " & _
            "'" & UpdtTag & "', '" & txtCVchk.Text & "');"
        MyID = GetFieldValue("SELECT CASHFLOWMGR.MGRCFID From CashFlowMgr WHERE CASHFLOWMGR.LASTUPDATE='" & UpdtTag & "';")
        txtCFid.Text = MyID
    Else
        LYRcon.Execute "Update CashFlowMgr SET BILLS = '" & tmpBill & "', EXPENSECV = '" & txtCVNO.Text & "', " & _
            "EXPENSEAMT = " & Val(Replace(txtCVamt.Text, ",", "")) & ", COINS = '" & tmpCoin & "', CFDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "', " & _
            "REMARKS = '" & txtRemarks.Text & "', PREPAREDBY = '" & cboCashier.Text & "', APPROVEDBY = '" & txtAprvby.Text & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & UpdtTag & "', PCVCHECK = '" & txtCVchk.Text & "' " & _
            "Where CASHFLOWMGR.MGRCFID = " & Val(txtCFid.Text) & " And CashFlowMgr.LASTUPDATE = '" & UpdtTag & "';"
    End If

    With flxChk
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 4) = "PDC DUE" Or .TextMatrix(i, 4) = "CURRENT" Then
                
                'LYRcon.Execute "UPDATE checks SET MgrCFid = " & MyID & " Where checks.CHECKPRID = " & Val(.TextMatrix(i, 7)) & ";"
                LYRcon.Execute "UPDATE  checks SET  MGRCFID = " & MyID & " Where checks.CHKSPRID = " & Val(.TextMatrix(i, 7)) & ";"
            ElseIf .TextMatrix(i, 4) = "ENCASHMENT" And Val(.TextMatrix(i, 7)) > 0 Then
                'LYRcon.Execute "UPDATE checks SET MgrCFid = " & MyID & " Where checks.CHECKPRID = " & Val(.TextMatrix(i, 7)) & ";"
                LYRcon.Execute "UPDATE checks SET MgrCFid = " & MyID & " Where checks.CHKSPRID = " & Val(.TextMatrix(i, 7)) & ";"
                
            ElseIf .TextMatrix(i, 4) = "ENCASHMENT" And Val(.TextMatrix(i, 7)) = 0 Then
                LYRcon.Execute "INSERT INTO checks(CHKACCTNAME, BANK, CHECKNO, AMOUNT, `STATUS`, CHECKDATE, DATE4BRKDWN, UPDATEDBY, LASTUPDATE, MGRCFID) " & _
                    "VALUES('" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 5) & "', '" & .TextMatrix(i, 1) & "', " & Val(Replace(.TextMatrix(i, 6), ",", "")) & ", 'ENCASHMENT', " & _
                    "'" & Format(DateValue(.TextMatrix(i, 2)), "yyyy-mm-dd") & "',  '" & Format$(CFdate.Value, "yyyy-mm-dd") & "', " & cboCashier.ItemData(cboCashier.ListIndex) & ", '" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "', " & MyID & ");"
            End If
        Next i
    End With
    
    'Save Received checks
    If flxRC.Rows > 0 Then
        With flxRC
            For i = 0 To .Rows - 1
                LYRcon.Execute "UPDATE Checks SET DATE4BRKDWN = '" & Format(CFdate.Value, "yyyy-mm-dd") & "', MGRCFID = " & MyID & " Where checks.CHKSPRID = " & Val(flxRC.TextMatrix(i, 7)) & ";"
            Next i
        End With
    End If

End Sub



Private Sub RXResizer1_ResizeComplete()
    Label3(4).Left = (Me.Width - 100) - Label3(4).Width
    Label3(5).Left = (Me.Width - 130) - Label3(5).Width
    With flxChk
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.15
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.25
        .ColWidth(4) = (.Width - 300) * 0.1
        .ColWidth(5) = (.Width - 300) * 0.2
        .ColWidth(6) = (.Width - 300) * 0.15
        .ColWidth(7) = 0
    End With
End Sub

Private Sub txtBill_Change(Index As Integer)
    Dim Ans As Single
    
    Ans = Val(txtBill(Index).Text) * Val(lblBill(Index).Caption)
    lblBillTtl(Index).Caption = Format$(Ans, "#,##0.00")
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

Private Sub txtChk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtChk.Tag = "skiplostfocus"
        txtChk.Visible = False
        flxChk.SetFocus
        flxChk.Col = 1
    End If
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

    If txtChk.Tag = "skiplostfocus" Then
        txtChk.Tag = ""
        Exit Sub
    End If
    If ActCol = 2 And txtChk.Text <> "" And Not IsDate(txtChk.Text) Then
        MsgBox "Invalid entry on Check Date.", vbExclamation + vbOKOnly, "Validation Error"
        txtChk.SetFocus
        txtChk.SelStart = 0
        txtChk.SelLength = Len(txtChk.Text)
        Exit Sub
    End If
    flxChk.TextMatrix(ActRow, ActCol) = txtChk.Text
    If ActCol = 1 And txtChk.Text <> "" Then flxChk.TextMatrix(ActRow, 4) = "ENCASHMENT"
    txtChk.Visible = False
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


Private Sub CFpreview()
    
    Dim txtH As TextObject
    Dim txtD As TextObject
    Dim lnObjD As LineObject
    Dim PageNumber As FieldObject
    
    Set M_Proj = New CRAXDRT.Application
    Set CF_Prvw = M_Proj.NewReport
    Set CF_Prvw = M_Proj.NewReport
    
    Dim Qry As String, BillBrk As String, CoinBrk As String, tofCV As String, CrntORseries As String
    'Dim NetCash As Currency, ChkTtlAmt As Currency, BillsTtl As Currency, CoinsTtl As Currency, Cash4Deposit As Currency, CollectorAmt As Currency, TtlPDCamt As Currency
    Dim i As Integer, Ln As Integer, cfRow As Integer, Hx As Integer, RowTtlColl As Integer, RowTtlDisb As Integer, ChkRec As Integer
    Dim CountPDC As Integer, CountCurrent As Integer, CountEnCash As Integer, TtlCrntAmt As Currency, TtlEncAmt As Double, CrntORamt As Currency, PdcORamt As Currency
    
    
    SSTab1.TabCaption(1) = "Processing Print Preview"
    DoEvents
    'Report settings
    CF_Prvw.PaperSize = crPaperLegal
    CF_Prvw.PaperOrientation = crPortrait
    CF_Prvw.TopMargin = 500
    CF_Prvw.PaperSize = crPaperLegal
    CF_Prvw.PaperOrientation = crPortrait
    
    'Section 1: Report Header
    CF_Prvw.Sections(1).Suppress = True
    'Section 2: Page Header
    CF_Prvw.Sections(2).Height = 1300
    Set txtH = CF_Prvw.Sections(2).AddTextObject(SysBranch, 500, 0)
    txtH.Width = 11000: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 12: txtH.Font.Bold = True: txtH.HorAlignment = crHorCenterAlign
    Qry = "SELECT branch.ADDRESS From branch Where branch.SYSTEMBRANCH = 1"
    Brn = GetFieldValue(Qry)
    Set txtH = CF_Prvw.Sections(2).AddTextObject(Brn, 500, 300)
    txtH.Width = 11000: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 10: txtH.Font.Bold = True: txtH.HorAlignment = crHorCenterAlign
    Set txtH = CF_Prvw.Sections(2).AddTextObject("MANAGER'S CASH FLOW", 500, 900)
    txtH.Width = 11000: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 12: txtH.Font.Bold = True: txtH.HorAlignment = crHorCenterAlign
    
    'section 3: Details
    CF_Prvw.Sections(3).Height = 13000
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Date :", 8000, 300)   'Date
    txtD.Width = 1000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(CFdate.Value, "Mmm. dd, yyyy"), 9000, 300)   'Date
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign: txtD.Font.Underline = True
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Undeposited Collections prior days", 500, 600)   'Undeposited Collections prior days
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
    If Val(txtUndpColl.Text) > 0 Then
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(Val(txtUndpColl.Text), "#,##0.00"), 9000, 600)  'Undeposited Collections prior days
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, 900, 10600, 900) 'HLine Undeposited Collections prior days
    lnObjD.LineThickness = 1
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("COLLECTION:", 500, 900)   'RECEIPTS
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.Font.Bold = True: txtD.HorAlignment = crLeftAlign
    Set txtD = CF_Prvw.Sections(3).AddTextObject("SparePart Cash [Walk-in]", 1500, 1300)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL AND (payments.REMARKS NOT LIKE '%WHOLESALE%' OR payments.REMARKS IS NULL);"
    SPcash = Val(GetFieldValue(Qry))
    If SPcash > 0 Then 'Show SP Cash Sales Amount
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(SPcash, "#,##0.00"), 7000, 1300)  'Cash Sales
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, 1580, 8600, 1580) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("SparePart Cash [Wholesale]", 1500, 1600)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL AND payments.REMARKS LIKE '%WHOLESALE%';"
    SPcashWS = Val(GetFieldValue(Qry))
    If SPcashWS > 0 Then 'Show SP Cash WholeSale Amount
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(SPcashWS, "#,##0.00"), 7000, 1600)  'Cash Sales
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, 1880, 8600, 1880) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Service", 1500, 1900)   'Service
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(payments.SERVICE) AS SumOfSvc FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL AND payments.CUSTOMERID = 0;"
    SvcAmt = Val(GetFieldValue(Qry))
    If SvcAmt > 0 Then 'Show Service Cash Sales Amount
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(SvcAmt, "#,##0.00"), 7000, 1900)  'Service Cash
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, 2180, 8600, 2180) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Down Payment", 1500, 2200)   'DownPayment
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(payments.DOWNPAYMENT) As SumOfDP From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL;"
    DPamt = Val(GetFieldValue(Qry))
    If DPamt > 0 Then 'Show DownPayment Amount
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(DPamt, "#,##0.00"), 7000, 2200)  'Down Payment
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, 2480, 8600, 2480) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Others", 1500, 2500)   'Others
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(payments.OTHERS) AS SumOfSvc FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL AND  payments.CUSTOMERID = 0;"
    OtherAmt = Val(GetFieldValue(Qry))
    If OtherAmt > 0 Then 'Show Other Amount
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(OtherAmt, "#,##0.00"), 7000, 2500)  'Others
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, 2780, 8600, 2780) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Collected A/R", 1500, 2800)   'Collected A/R
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Spare Parts", 2000, 3100)    'SP Colltd AR
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(PAYMENTS.SPAREPARTS) As SumOfSP From PAYMENTS Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL;"
    SPonAcct = Val(GetFieldValue(Qry))
    If SPonAcct > 0 Then 'Show SP On Account
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(SPonAcct, "#,##0.00"), 5000, 3100)  'SP OnAcct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(5000, 3380, 6600, 3380) 'HLine SP On Acct
    lnObjD.LineThickness = 1
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Service On-Account", 2000, 3400)    'Service REG  AR
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(payments.SERVICE) AS SumOfSvc FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL AND payments.CUSTOMERID > 0;"
    SvcReg = Val(GetFieldValue(Qry))
    If SvcReg > 0 Then 'Service REG A/R
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(SvcReg, "#,##0.00"), 5000, 3400)  'SP OnAcct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(5000, 3680, 6600, 3680) 'HLine SvcReg On Acct
    lnObjD.LineThickness = 1
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Service FSC", 2000, 3700)    'Service FSC  AR
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    'Qry = "SELECT SUM(PAYMENTS.SPAREPARTS) As SumOfSP From PAYMENTS Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.ORNO IS NOT NULL AND PAYMENTS.EmpId = " & cboCashier.ItemData(cboCashier.ListIndex) & ";"
    'SvcFSC = Val(GetFieldValue(Qry))
    If SvcFSC > 0 Then 'Service REG A/R
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(SvcFSC, "#,##0.00"), 5000, 3700)  'SP OnAcct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(5000, 3980, 6600, 3980) 'HLine SvcFSC On Acct
    lnObjD.LineThickness = 1
    
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Others", 2000, 4000)    'Others  AR
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(payments.OTHERS) AS SumOfSvc FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL AND payments.CUSTOMERID > 0;"
    OthrAcct = Val(GetFieldValue(Qry))
    If OthrAcct > 0 Then 'Others A/R
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(OthrAcct, "#,##0.00"), 5000, 4000)  'SP OnAcct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(5000, 4280, 6600, 4280) 'HLine SvcFSC On Acct
    lnObjD.LineThickness = 1
    
    'Display Total A/R
    TtlAR = OthrAcct + SvcFSC + SvcReg + SPonAcct
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlAR, "#,##0.00"), 7000, 4000) 'Total Collection
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, 4280, 8600, 4280) 'HLine COllector
    lnObjD.LineThickness = 1
       
    'Display Add: Cash Returned - Disburser
    cfRow = 4000
    Qry = "SELECT SUM(cashreturn.CRAMOUNT) As sumCRAMOUNT From cashreturn " & _
          "Where cashreturn.CRDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND cashreturn.PRINTCOPY > 0 AND cashreturn.REMARKS NOT LIKE '%CANCELLED%' ;"
    CashReturn = Val(GetFieldValue(Qry))
    If CashReturn > 0 Then 'Show Cash Return Amount
        cfRow = cfRow + 300
        Set txtD = CF_Prvw.Sections(3).AddTextObject("Add: Cash Returned", 1500, cfRow) 'Add: Cash Returned
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(CashReturn, "#,##0.00"), 9000, cfRow) 'cash return Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine cashreturn
        lnObjD.LineThickness = 1
    End If
    If Val(Replace(txtRecTOF.Text, ",", "")) > 0 Then
        cfRow = cfRow + 300
        If Val(Replace(txtRecTOF.Text, ",", "")) > 0 Then  'display in column 2000
            Set txtD = CF_Prvw.Sections(3).AddTextObject("Received Funds", 2000, cfRow) 'Add: Received funds from TOF
            txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Else
            Set txtD = CF_Prvw.Sections(3).AddTextObject("Add: Received Funds", 1500, cfRow) 'Add: Received funds from TOF
            txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        End If
        Set txtD = CF_Prvw.Sections(3).AddTextObject(txtRecTOF.Text, 5000, cfRow) 'Ttl Rec. TOF
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(5000, cfRow + 280, 6600, cfRow + 280) 'HLine cashreturn
        lnObjD.LineThickness = 1                '5000, 4280, 6600, 4280
    End If
    
    'Display Total Cash Return and Received funds
    If CashReturn > 0 Or Val(Replace(txtRecTOF.Text, ",", "")) > 0 Then
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(CashReturn + Val(Replace(txtRecTOF.Text, ",", "")), "#,##0.00"), 7000, cfRow) 'Total Collection
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine COllector
        lnObjD.LineThickness = 1
    End If
    
    'Display TOTAL COLLECTION
    TtlColltn = SPcash + SPcashWS + SvcAmt + DPamt + OtherAmt + TtlAR + CashReturn + Val(Replace(txtRecTOF.Text, ",", ""))
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlColltn, "#,##0.00"), 9000, cfRow) 'TOTAL COLLECTION
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine TOTALRECEIPTS
    lnObjD.LineThickness = 1
    
    
    'Display Less: DISBURSEMENT
    cfRow = cfRow + 500
    Set txtD = CF_Prvw.Sections(3).AddTextObject("LESS:", 500, cfRow) 'LESS
    txtD.Width = 1000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set txtD = CF_Prvw.Sections(3).AddTextObject("DISBURSEMENTS", 1500, cfRow) 'DISBURSEMENT
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    'Display Advances for Liquidation
    cfRow = cfRow + 300
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Advances for Liquidation - Disburser", 1500, cfRow) 'Advances for Liquidation - Disburser
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(advliquidation.AFLAMOUNT) As sumAFLamt From advliquidation " & _
          "Where advliquidation.AFLDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND  advliquidation.PRINTCOPY > 0;"
    AdvLiq = Val(GetFieldValue(Qry))
    If AdvLiq > 0 Then 'Show Cash Return Amount
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(AdvLiq, "#,##0.00"), 7000, cfRow) 'Adv. Liquidation
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Else    'View Total of CV if no adv. for liquidation available
        If CFdate.Value <= DateValue("5/11/2010") Then 'Date of Upgrade
            Qry = "SELECT CashVoucher.AMOUNT From CashVoucher Where cashvoucher.CVDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND cashvoucher.REMARKS = 'EXPENSE' AND cashvoucher.PRINTCOPY >= 1"
            CVamt = Val(GetFieldValue(Qry))
            If CVamt > 0 Then 'Show Cash Return Amount
                Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(CVamt, "#,##0.00"), 7000, cfRow) 'Adv. Liquidation
                txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
            End If
        End If
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Adv. Liquidation
    lnObjD.LineThickness = 1
    
    'Display Cash out
    cfRow = cfRow + 300
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Cash Out", 1500, cfRow) 'Cash Out
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Qry = "SELECT SUM(CashVoucher.AMOUNT) As SumOfAMT From CashVoucher Where cashvoucher.CVDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND cashvoucher.REMARKS <> 'EXPENSE' AND cashvoucher.REMARKS <> 'TOF' AND cashvoucher.REMARKS <> 'CANCELLED';"
    CashOutAmt = Val(GetFieldValue(Qry))
    If CashOutAmt > 0 Then 'Show Cash Out Amount
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(CashOutAmt, "#,##0.00"), 7000, cfRow) 'Cash Out Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine cash out
    lnObjD.LineThickness = 1
    
    
    
    'Display Sales Return
    Qry = "SELECT SUM(sreturnsp.TOTALAMT) As SumOfTtlAmt From sreturnsp INNER JOIN spinvoice ON (sreturnsp.INVOICE = spinvoice.SPINVNO) " & _
          "Where sreturnsp.RETDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND sreturnsp.`PROCESS` = 'REPLACE' AND Case When sreturnsp.INVTYPE = 0 Then spinvoice.INVTYPE = 'CREDIT' Else spinvoice.INVTYPE = 'CASH' End"
    sRetAmt = Val(GetFieldValue(Qry))
    If sRetAmt > 0 Then 'Show Sales Return Amount
        cfRow = cfRow + 300
        Set txtD = CF_Prvw.Sections(3).AddTextObject("Sales Return", 1500, cfRow) 'Less Online Payment
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(sRetAmt, "#,##0.00"), 7000, cfRow) 'Sales REturn Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine ONLINE PAYMENT
        lnObjD.LineThickness = 1
    End If
   
    
    'Display LESS: ONLINE PAYMENT
    'Service
    Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
            "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND (PAYMENTS.PAYTYPE = 'OLP' OR PAYMENTS.REMARKS LIKE '%ON-LINE%');"
    OnLineSERVICEAmt = Val(GetFieldValue(Qry))
    'Others
    Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
            "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND (PAYMENTS.PAYTYPE = 'OLP' OR PAYMENTS.REMARKS LIKE '%ON-LINE%');"
    OnLineOthersAmt = Val(GetFieldValue(Qry))
    'SP
    Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
            "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND (PAYMENTS.PAYTYPE = 'OLP' OR PAYMENTS.REMARKS LIKE '%ON-LINE%');"
    OnLineSPAmt = Val(GetFieldValue(Qry))
    
    OnLineAmt = OnLineSERVICEAmt + OnLineOthersAmt + OnLineSPAmt
    If OnLineAmt > 0 Then 'Show Online Amount
        cfRow = cfRow + 300
        Set txtD = CF_Prvw.Sections(3).AddTextObject("Online Payment", 1500, cfRow) 'Less Online Payment
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(OnLineAmt, "#,##0.00"), 7000, cfRow) 'ONLINE Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine ONLINE PAYMENT
        lnObjD.LineThickness = 1
    End If
         
    'Display LESS: Credit Card PAYMENT
    Qry = "SELECT SUM(payments.DOWNPAYMENT + payments.SPAREPARTS + payments.SERVICE + payments.Others) As SumOfCCard " & _
            "FROM payments WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS LIKE '%CREDIT CARD%';"
    CCardAmt = Val(GetFieldValue(Qry))
    If CCardAmt > 0 Then 'Show Creditcard Amount
        cfRow = cfRow + 300
        Set txtD = CF_Prvw.Sections(3).AddTextObject("Credit Card Payment", 1500, cfRow) 'Less Online Payment
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(CCardAmt, "#,##0.00"), 7000, cfRow) 'ONLINE Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine ONLINE PAYMENT
        lnObjD.LineThickness = 1
    End If
    
    'Display LESS: Inter-Branch PAYMENT
    Qry = "SELECT SUM(payments.DOWNPAYMENT + payments.SPAREPARTS + payments.SERVICE + payments.Others) As SumOfInterBrn " & _
          "From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.REMARKS LIKE '%INTER-BRANCH%'"
    IBrnAmt = Val(GetFieldValue(Qry))
    If IBrnAmt > 0 Then 'Show Inter Branch Amount
        cfRow = cfRow + 300
        Set txtD = CF_Prvw.Sections(3).AddTextObject("Inter-Branch Payment", 1500, cfRow) 'Less Online Payment
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(IBrnAmt, "#,##0.00"), 7000, cfRow) 'ONLINE Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine ONLINE PAYMENT
        lnObjD.LineThickness = 1
    End If
    TtlLess = AdvLiq + CashOutAmt + sRetAmt + OnLineAmt + CCardAmt + IBrnAmt
   
    'TtlLess = AdvLiq + CashOutAmt + sRetAmt + OnLineAmt + CCardAmt + IBrnAmt
    'Display TOTAL of DISBURSEMENTS
    If TtlLess > 0 Then
        Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlLess, "#,##0.00"), 9000, cfRow) 'TOTAL DISBURSEMENT
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine TOTAL DISBURSEMENT
    lnObjD.LineThickness = 1
    
    'Display NET CASH ON HAND
    cfRow = cfRow + 500
    Set txtD = CF_Prvw.Sections(3).AddTextObject("NET COLLECTION", 500, cfRow) 'NET CASH ON HAND
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    NetColl = TtlColltn - TtlLess
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(NetColl, "#,##0.00"), 9000, cfRow) 'NET CASH ON HAND
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine NET CASH ON HAND
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = 2 'crLSDoubleLine
    
    'Display BREAKDOWN
    cfRow = cfRow + 500
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Breakdown:", 500, cfRow) 'Breakdown
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    
    'Display PDC CHECKs issued with OR
    cfRow = cfRow + 300
    ChkTtlAmt = 0
    TtlPDCamt = 0
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Post-Dated Checks Due", 1500, cfRow) 'Breakdown
    txtD.Width = 5000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    CountPDC = CountCheck("PDC DUE")
    If CountPDC > 0 Then
        If CountPDC > 2 Then 'Compute total for Checks add Display Check in Attachment
            'For i = 1 To flxChk.Rows - 1
            '    TtlPDCamt = TtlPDCamt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
            'Next i
            For i = 1 To flxChk.Rows - 1
                If flxChk.TextMatrix(i, 4) = "PDC DUE" Then
                    TtlPDCamt = TtlPDCamt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
                End If
            Next i
            
            cfRow = cfRow + 300
            Set txtD = CF_Prvw.Sections(3).AddTextObject(CountPDC & " PDC (Please see Attachment)", 2000, cfRow) 'Total of All Checks (Summary)
            txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
            Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlPDCamt, "#,##0.00"), 9000, cfRow) 'Total Check Amount
            txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
        Else    'Display All Checks (limited to 3)
            For i = 0 To flxChk.Rows - 1
                If flxChk.TextMatrix(i, 4) = "PDC DUE" Then
                    cfRow = cfRow + 300
                    Set txtD = CF_Prvw.Sections(3).AddTextObject("Check# " & flxChk.TextMatrix(i, 1) & " - " & flxChk.TextMatrix(i, 3) & ")", 2000, cfRow) 'Checks NO and Acct Name
                    txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
                    Set txtD = CF_Prvw.Sections(3).AddTextObject(flxChk.TextMatrix(i, 6), 7000, cfRow) 'Check Amount
                    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
                    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 6800, cfRow + 280) 'HLine Checks
                    lnObjD.LineThickness = 1
                    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
                    lnObjD.LineThickness = 1
                    TtlPDCamt = TtlPDCamt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
                End If
            Next i
            'Display Check Total Amount
            Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlPDCamt, "#,##0.00"), 9000, cfRow) 'Total PDC Amount
            txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
        End If
    Else
        cfRow = cfRow + 300
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 6800, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    End If
    ChkTtlAmt = TtlPDCamt
    
    'Display Current CHECKs
    cfRow = cfRow + 300
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Current Checks:", 1500, cfRow) 'Breakdown
    txtD.Width = 2500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    CountCurrent = CountCheck("CURRENT")
    If CountCurrent > 0 Then
        If CountCurrent > 2 Then 'Compute total for Checks add Display Check in Attachment
            'For i = 1 To flxChk.Rows - 1
            '    TtlCrntAmt = TtlCrntAmt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
            'Next i
            For i = 1 To flxChk.Rows - 1
                If flxChk.TextMatrix(i, 4) = "CURRENT" Then
                    TtlCrntAmt = TtlCrntAmt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
                End If
            Next i
            
            cfRow = cfRow + 300
            Set txtD = CF_Prvw.Sections(3).AddTextObject(CountCurrent & " Current Checks (Please see Attachment)", 2000, cfRow) 'Total of All Current Checks (Summary)
            txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
            Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlCrntAmt, "#,##0.00"), 9000, cfRow) 'Total Check Amount
            txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
        Else    'Display All Checks (limited to 3)
            For i = 0 To flxChk.Rows - 1
                If flxChk.TextMatrix(i, 4) = "CURRENT" Then
                    cfRow = cfRow + 300
                    Set txtD = CF_Prvw.Sections(3).AddTextObject("Check# " & flxChk.TextMatrix(i, 1) & " - " & flxChk.TextMatrix(i, 3) & ")", 2000, cfRow) 'Checks NO and Acct Name
                    txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
                    Set txtD = CF_Prvw.Sections(3).AddTextObject(flxChk.TextMatrix(i, 6), 7000, cfRow) 'Check Amount
                    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
                    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 6800, cfRow + 280) 'HLine Checks
                    lnObjD.LineThickness = 1
                    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
                    lnObjD.LineThickness = 1
                    TtlCrntAmt = TtlCrntAmt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
                End If
            Next i
            'Display Check Total Amount
            Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlCrntAmt, "#,##0.00"), 9000, cfRow) 'Total PDC Amount
            txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
        End If
    Else
        cfRow = cfRow + 300
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 7000, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    End If
    ChkTtlAmt = ChkTtlAmt + TtlCrntAmt
    'Display ENCASHMENT CHECKs
    cfRow = cfRow + 300
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Check Encashment:", 1500, cfRow) 'Breakdown
    txtD.Width = 2500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    TtlEncAmt = 0
    CountEnCash = CountCheck("ENCASHMENT")
    If CountEnCash > 0 Then
        If CountEnCash > 2 Then 'Compute total for Checks add Display Check in Attachment
            'For i = 1 To flxChk.Rows - 1
            '    TtlEncAmt = TtlEncAmt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
            'Next i
            For i = 1 To flxChk.Rows - 1
                If flxChk.TextMatrix(i, 4) = "ENCASHMENT" Then
                    TtlEncAmt = TtlEncAmt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
                End If
            Next i
            cfRow = cfRow + 300
            Set txtD = CF_Prvw.Sections(3).AddTextObject(CountEnCash & " Check Encashment (Please see Attachment)", 2000, cfRow) 'Total of All Current Checks (Summary)
            txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
            Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlEncAmt, "#,##0.00"), 9000, cfRow) 'Total Check Amount
            txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
        Else    'Display All Checks (limited to 3)
            For i = 0 To flxChk.Rows - 1
                If flxChk.TextMatrix(i, 4) = "ENCASHMENT" Then
                    cfRow = cfRow + 300
                    Set txtD = CF_Prvw.Sections(3).AddTextObject("Check# " & flxChk.TextMatrix(i, 1) & " - " & flxChk.TextMatrix(i, 3), 2000, cfRow) 'Checks NO and Acct Name
                    txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
                    Set txtD = CF_Prvw.Sections(3).AddTextObject(flxChk.TextMatrix(i, 6), 7000, cfRow) 'Check Amount
                    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
                    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 6600, cfRow + 280) 'HLine Checks
                    lnObjD.LineThickness = 1
                    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
                    lnObjD.LineThickness = 1
                    TtlEncAmt = TtlEncAmt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
                End If
            Next i
            'Display Check Total Amount
            Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(TtlEncAmt, "#,##0.00"), 9000, cfRow) 'Total PDC Amount
            txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
            Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
            lnObjD.LineThickness = 1
        End If
    Else
        cfRow = cfRow + 300
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 7000, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    End If
    ChkTtlAmt = ChkTtlAmt + TtlEncAmt
    
    'Display Bills
    cfRow = cfRow + 300
    BillsTtl = 0
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Bills:", 1500, cfRow) 'Breakdown
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    For i = 0 To 6
        cfRow = cfRow + 300
        BillBrk = Space$(6 - Len(txtBill(i).Text)) & txtBill(i).Text & Space$(5) & "X" & Space$(10 - Len(lblBill(i).Caption)) & lblBill(i).Caption
        Set txtD = CF_Prvw.Sections(3).AddTextObject(BillBrk, 2000, cfRow) 'Breakdown Bills
        txtD.Width = 3300: txtD.Height = 300: txtD.Font.Name = "COURIER NEW": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject(lblBillTtl(i).Caption, 7000, cfRow) 'BILLS Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine BILL AMT
        lnObjD.LineThickness = 1
        BillsTtl = BillsTtl + Val(Replace(lblBillTtl(i).Caption, ",", ""))
    Next i
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(BillsTtl, "#,##0.00"), 9000, cfRow) 'Total BILLS Amount
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
    lnObjD.LineThickness = 1
    'Display Coins
    cfRow = cfRow + 300
    CoinsTtl = 0
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Coins:", 1500, cfRow) 'Breakdown Coins
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    For i = 0 To 5  'Coins Breakdown
        cfRow = cfRow + 300
        CoinBrk = Space$(6 - Len(txtCoin(i).Text)) & txtCoin(i).Text & Space$(5) & "X" & Space$(10 - Len(lblCoin(i).Caption)) & lblCoin(i).Caption
        Set txtD = CF_Prvw.Sections(3).AddTextObject(CoinBrk, 2000, cfRow) 'Breakdown coins
        txtD.Width = 3300: txtD.Height = 300: txtD.Font.Name = "COURIER NEW": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject(lblCoinTtl(i).Caption, 7000, cfRow) 'coins Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Coins AMT
        lnObjD.LineThickness = 1
        CoinsTtl = CoinsTtl + Val(Replace(lblCoinTtl(i).Caption, ",", ""))
    Next i
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(CoinsTtl, "#,##0.00"), 9000, cfRow) 'Total CoinS Amount
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
    lnObjD.LineThickness = 1
    'Display CASH ON HAND FOR DEPOSIT
    cfRow = cfRow + 400
    Set txtD = CF_Prvw.Sections(3).AddTextObject("T O T A L", 500, cfRow) 'CASH ON HAND FOR DEPOSIT
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    Cash4Deposit = ChkTtlAmt + BillsTtl + CoinsTtl
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(Cash4Deposit, "#,##0.00"), 9000, cfRow) 'Total CoinS Amount
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = 2
    'Display Less: DEPOSIT
    cfRow = cfRow + 400
    Set txtD = CF_Prvw.Sections(3).AddTextObject("ADD: " & IIf(txtCVNO.Text <> "", "CV# " & txtCVNO.Text, "") & IIf(txtCVchk.Text <> "", " [CHECK " & txtCVchk.Text & "]", ""), 500, cfRow) 'CASH ON HAND FOR DEPOSIT
    txtD.Width = 6000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(Val(Replace(txtCVamt.Text, ",", "")), "#,##0.00"), 9000, cfRow) 'CHECK VOUCHER Amount
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
    lnObjD.LineThickness = 1
    'Display Cash Balance, End
    cfRow = cfRow + 400
    Set txtD = CF_Prvw.Sections(3).AddTextObject("TOTAL AMOUNT FOR DEPOSIT", 500, cfRow) 'Cash Balance, End
    txtD.Width = 6000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    Set txtD = CF_Prvw.Sections(3).AddTextObject(Format(Cash4Deposit + Val(Replace(txtCVamt.Text, ",", "")), "#,##0.00"), 9000, cfRow) 'CHECK VOUCHER Amount
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = crLSDoubleLine
    'Display Remarks
    cfRow = cfRow + 400
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Remarks:", 500, cfRow) 'Remarks
    txtD.Width = 6000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    If NetColl - Cash4Deposit > 0 Then
        Set txtD = CF_Prvw.Sections(3).AddTextObject("Short: " & Format(Abs(NetColl - Cash4Deposit), "#,##0.00"), 2000, cfRow) 'Remarks
        txtD.Width = 6000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    ElseIf NetColl - Cash4Deposit < 0 Then
        Set txtD = CF_Prvw.Sections(3).AddTextObject("Excess: " & Format(Abs(NetColl - Cash4Deposit), "#,##0.00"), 2000, cfRow) 'Remarks
        txtD.Width = 6000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    End If
    'Encoded Remarks
    If txtRemarks.Text <> "" Then
        cfRow = cfRow + 300
        Set txtD = CF_Prvw.Sections(3).AddTextObject(txtRemarks.Text, 2000, cfRow)  'Remarks
        txtD.Width = 10000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = False
    End If
    'OR Issued to PDC
    With flxChk
        If CountPDC > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 4) = "PDC DUE" Then
                    Qry = "SELECT checkalloc.ORNO FROM checkalloc " & _
                          "WHERE  checkalloc.ORDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 7)) & " ORDER BY checkalloc.ORNO"
                    PDCor = GetOR(Qry)
                    Qry = "SELECT SUM(checkalloc.ORAmt) As SumOfORAmt From checkalloc " & _
                          "Where checkalloc.ORDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 7)) & ";"
                    PdcORamt = Val(GetFieldValue(Qry))
                    If PdcORamt > 0 Then
                        cfRow = cfRow + 250
                        Set txtD = CF_Prvw.Sections(3).AddTextObject("PDC Due Check# " & .TextMatrix(i, 1) & " [Receipt Amt: " & Format(PdcORamt, "#,##0.00") & " Receipt#: " & PDCor & "]", 1000, cfRow)   'OR Series PDC
                        txtD.Width = 9500: txtD.Height = 250: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
                    End If
                End If
            Next i
        End If
        
        If CountCurrent > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 4) = "CURRENT" Then
                    Qry = "SELECT checkalloc.ORNO FROM checkalloc " & _
                          "WHERE  checkalloc.ORDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 7)) & " ORDER BY checkalloc.ORNO"
                    CrntORseries = GetOR(Qry)
                    Qry = "SELECT SUM(checkalloc.ORAmt) As SumOfORAmt From checkalloc " & _
                          "Where checkalloc.ORDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 7)) & ";"
                    CrntORamt = Val(GetFieldValue(Qry))
                    If CrntORamt > 0 Then
                        cfRow = cfRow + 250
                        Set txtD = CF_Prvw.Sections(3).AddTextObject("Receipt issued to Current Check# " & .TextMatrix(i, 1) & " [Receipt Amt: " & Format(CrntORamt, "#,##0.00") & " Receipt#: " & CrntORseries & "]", 1000, cfRow)   'OR Series PDC
                        txtD.Width = 9500: txtD.Height = 250: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign
                    End If
                End If
            Next i
        End If
    End With
                
    'Checks Received
    If flxRC.Rows >= 1 Then
        If flxRC.Rows <= 2 Then
            For i = 0 To flxRC.Rows - 1
                cfRow = cfRow + 250
                Set txtD = CF_Prvw.Sections(3).AddTextObject("Check received from " & flxRC.TextMatrix(i, 3) & " Check#: " & flxRC.TextMatrix(i, 1) & " Date: " & Format(flxRC.TextMatrix(i, 2)) & " Amount: " & flxRC.TextMatrix(i, 6), 1000, cfRow) 'Checks received
                txtD.Width = 9500: txtD.Height = 250: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
            Next i
        Else    'Summarize
            cfRow = cfRow + 200
            Set txtD = CF_Prvw.Sections(3).AddTextObject(flxRC.Rows & " checks received. Please see attachment for details.", 1000, cfRow)  'Checks received
            txtD.Width = 9500: txtD.Height = 250: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        End If
    End If
    
    If cfRow > 14500 Then
        cfRow = cfRow + 500
    Else
        cfRow = 15200
    End If
    'Display Cash Balance, End
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Prepared by:", 500, cfRow)    '15200)
    txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set txtD = CF_Prvw.Sections(3).AddTextObject("Noted by:", 6500, cfRow)   '15200)
    txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    cfRow = cfRow + 400
    Set txtD = CF_Prvw.Sections(3).AddTextObject(UCase$(cboCashier.Text), 2000, cfRow)  '15600)
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
    Set txtD = CF_Prvw.Sections(3).AddTextObject(UCase$(txtAprvby.Text), 7500, cfRow)   '15600)
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(2000, cfRow + 280, 5000, cfRow + 280) '15880)
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = crLSDoubleLine
    Set lnObjD = CF_Prvw.Sections(3).AddLineObject(7500, cfRow + 280, 10500, cfRow + 280)  '15880)
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = crLSDoubleLine
    cfRow = cfRow + 300
' ******************** CODE UPDATED: 06/08/2017 *************************
    If SysBranch = "BINOS SPAREPARTS STO. TOMAS" Or SysBranch = "BINOS SPAREPARTS SAMAL" Or SysBranch = "BINOS SPAREPARTS POLOMOLOK" Or SysBranch = "BINOS SPAREPARTS TAGUM" Or SysBranch = "MOTORHUB SPAREPARTS TAGUM" Or SysBranch = "MOTORHUB SPAREPARTS MARBEL" Then
    'diri mag'edit
        Set txtD = CF_Prvw.Sections(3).AddTextObject("SPARE PARTS OIC", 2000, cfRow)    '15900)
        txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject("BRANCH OIC", 7500, cfRow)    '15900)
        txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
    Else
        Set txtD = CF_Prvw.Sections(3).AddTextObject("MANAGER", 2000, cfRow)    '15900)
        txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject("AOM", 7500, cfRow)    '15900)
        txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
    End If
' **************** END CODE UPDATED: 06/08/2017 *************************
       
    'Display Report in Viewer
    CRViewer1.ReportSource = CF_Prvw
    CRViewer1.DisplayGroupTree = False
    CRViewer1.ViewReport
    CRViewer1.Zoom 100
    
    SSTab1.TabCaption(1) = "Print Preview"
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    
    
    Call CFprnAttach
        
End Sub


Private Sub CFprnAttach()
    Dim txtH As TextObject
    Dim txtD As TextObject
    Dim lnObjD As LineObject
    Dim PageNumber As FieldObject
    
    Set M_Proj = New CRAXDRT.Application
    Set CF_Attc = M_Proj.NewReport
    Set CF_Attc = M_Proj.NewReport
    
    Dim Qry As String, tofCV As String, CrntORseries As String
    Dim ORCashSales As String, ORCollCashier As String, ORMotorShop As String, ORSvc As String, ORDown As String
    Dim CCardOR As String, OnlineOR As String, PDCor As String, CashRetNO As String, AdvLiqNO As String, CVmfkg As String, CVpers As String, CVtof As String
    Dim i As Integer, Ln As Integer, cfRow As Long, Hx As Integer, RowTtlColl As Integer, RowTtlDisb As Integer, ChkRec As Integer
    'Dim NetCash As Currency, ChkTtlAmt As Currency, CoinsTtl As Currency, Cash4Deposit As Currency, CollectorAmt As Currency, TtlPDCamt As Currency
    Dim CountPDC As Integer, CountCurrent As Integer, CountEnCash As Integer, TtlCrntAmt As Currency, TtlEncAmt As Double, CrntORamt As Currency
    Dim OnlineINV As String, IBrnINV As String, IBrnOR As String
    
    
    SSTab1.TabCaption(2) = "Processing Attachment"
    DoEvents
    
    'Report settings
    CF_Attc.PaperSize = crPaperLegal
    CF_Attc.PaperOrientation = crPortrait
    CF_Attc.TopMargin = 500
    'Section 1: Report Header
    CF_Attc.Sections(1).Suppress = True
    'Section 2: Page Header
    CF_Attc.Sections(2).Height = 900
    Set txtH = CF_Attc.Sections(2).AddTextObject(SysBranch, 500, 0)
    txtH.Width = 5000: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 12: txtH.Font.Bold = True: txtH.HorAlignment = crLeftAlign
    Set txtH = CF_Attc.Sections(2).AddTextObject(Brn, 500, 300)
    txtH.Width = 5000: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 10: txtH.Font.Bold = True: txtH.HorAlignment = crLeftAlign
    Set txtH = CF_Attc.Sections(2).AddTextObject("MANAGER'S CASH FLOW ATTACHMENT", 5500, 0)
    txtH.Width = 5500: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 12: txtH.Font.Bold = True: txtH.HorAlignment = crRightAlign
    'Set txtH = CF_Attc.Sections(2).AddTextObject("A T T A C H M E N T", 5000, 300)
    'txtH.Width = 5500: txtH.Height = 300: txtH.Font.Name = "ARIAL": txtH.Font.Size = 10: txtH.Font.Underline = True: txtH.HorAlignment = crRightAlign
    ' Add the page number to the footer
    Set PageNumber = CF_Attc.Sections(2).AddSpecialVarFieldObject(crSVTPageNofM, 10000, 300)
    PageNumber.Font.Size = 10
    PageNumber.Font.Name = "TAHOMA"
    
    'section 3: Details
    CF_Attc.Sections(3).Height = 14000  '14000
    Set txtD = CF_Attc.Sections(3).AddTextObject("Date :", 8000, 0)   'Date
    txtD.Width = 1000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
    Set txtD = CF_Attc.Sections(3).AddTextObject(Format(CFdate.Value, "Mmm. dd, yyyy"), 9000, 0)   'Date
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign: txtD.Font.Underline = True
    cfRow = 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("Undeposited Collections prior days", 500, cfRow)   'Undeposited Collections prior days
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crLeftAlign
    If Val(txtUndpColl.Text) > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(Val(txtUndpColl.Text), "P   #,##0.00"), 9000, cfRow)  'Undeposited Collections prior days
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10:  txtD.HorAlignment = crRightAlign
    End If
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Undeposited Collections prior days
    lnObjD.LineThickness = 1
    cfRow = cfRow + 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("COLLECTIONS:", 500, cfRow)   'RECEIPTS
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.Font.Bold = True: txtD.HorAlignment = crLeftAlign
    
    cfRow = cfRow + 300     'Walkin
    Set txtD = CF_Attc.Sections(3).AddTextObject("SpareParts Cash - Walkin", 1500, cfRow)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'if Cash Sales if available display OR Series
    If SPcash > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(SPcash, "P #,##0.00"), 7000, cfRow)  'Cash Sales Walkin
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        'Build Qry to display OR
        Qry = "SELECT payments.INVOICENO From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.SPAREPARTS > 0 AND payments.INVOICENO IS NOT NULL AND (payments.REMARKS NOT LIKE '%WHOLESALE%' OR payments.REMARKS IS NULL);"
        ORCashSales = GetOR(Qry)
        Hx = IIf(Len(ORCashSales) > 90, Round((Len(ORCashSales) / 90 + 0.49), 2), 1)
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject("[INV#: " & ORCashSales & "]", 2000, cfRow)   'Cash Sales OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (170 * Hx)
    End If
    
    cfRow = cfRow + 300     'WholeSale
    Set txtD = CF_Attc.Sections(3).AddTextObject("SpareParts Cash - Wholesale", 1500, cfRow)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'if Cash WholeSales if available display OR Series
    If SPcashWS > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(SPcashWS, "P #,##0.00"), 7000, cfRow)  'Cash Sales Walkin
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        'Build Qry to display OR
        Qry = "SELECT payments.INVOICENO From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.SPAREPARTS > 0 AND payments.INVOICENO IS NOT NULL AND payments.REMARKS LIKE '%WHOLESALE%';"
        ORCashWS = GetOR(Qry)
        Hx = IIf(Len(ORCashWS) > 0, Round((Len(ORCashWS) / 90 + 0.49), 2), 1)
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject("[INV#: " & ORCashWS & "]", 2000, cfRow)   'Cash WholeSale OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
    
    cfRow = cfRow + 300     'Service cash
    Set txtD = CF_Attc.Sections(3).AddTextObject("Service", 1500, cfRow)   'Service
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'if Service if available display OR Series
    If SvcAmt > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(SvcAmt, "P #,##0.00"), 7000, cfRow)  'Cash Sales Walkin
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        'Build Qry to display OR
        Qry = "SELECT payments.ORNO FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND payments.INVOICENO IS NULL AND  payments.CUSTOMERID = 0;"
        ORSvc = GetOR(Qry)
        cfRow = cfRow + 300
        Hx = Round((Len(ORSvc) / 90 + 0.99), 2)
        Set txtD = CF_Attc.Sections(3).AddTextObject("[OR#: " & ORSvc & "]", 2000, cfRow)   'Cash WholeSale OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
        
    cfRow = cfRow + 300     'Down Payment
    Set txtD = CF_Attc.Sections(3).AddTextObject("Down Payment", 1500, cfRow)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'if DownPayment is available display OR Series
    If DPamt > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(DPamt, "P #,##0.00"), 7000, cfRow)  'DP
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        'Build Qry to display OR
        Qry = "SELECT payments.ORNO From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.DOWNPAYMENT > 0 AND payments.INVOICENO IS NULL;"
        ORDown = GetOR(Qry)
        Qry = "SELECT payments.CRNO From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.DOWNPAYMENT > 0 AND payments.INVOICENO IS NULL;"
        CRseries = GetOR(Qry)
        If ORDown <> "" Then
            ORDown = "[OR#:" & ORDown & IIf(CRseries <> "", ", CR#:" & CRseries & "]", "]")
        Else
            ORDown = IIf(CRseries <> "", "[CR#:" & CRseries & "]", "")
        End If
        cfRow = cfRow + 300
        Hx = IIf(Len(ORDown) > 90, Round((Len(ORDown) / 90 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject(ORDown, 2000, cfRow)   'Down Payment OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
    
    cfRow = cfRow + 300     'Others
    Set txtD = CF_Attc.Sections(3).AddTextObject("Others", 1500, cfRow)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'if Others is available display OR Series
    If OtherAmt > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(OtherAmt, "P #,##0.00"), 7000, cfRow)  'Others
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Qry = "SELECT payments.ORNO FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.OTHERS > 0 AND payments.INVOICENO IS NULL AND payments.CUSTOMERID = 0;"
        ORothers = GetOR(Qry)
        Qry = "SELECT payments.CRNO FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.OTHERS > 0 AND payments.INVOICENO IS NULL AND payments.CUSTOMERID = 0;"
        CRseries = GetOR(Qry)
        If ORothers <> "" Then
            ORothers = "[OR#:" & ORothers & IIf(CRseries <> "", ", CR#:" & CRseries & "]", "]")
        Else
            ORothers = IIf(CRseries <> "", "[CR#:" & CRseries & "]", "")
        End If
        cfRow = cfRow + 300
        Hx = IIf(Len(ORothers) > 90, Round((Len(ORothers) / 90 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject(ORothers, 2000, cfRow)   'Down Payment OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If


    cfRow = cfRow + 300     'Collected A/R
    Set txtD = CF_Attc.Sections(3).AddTextObject("Collected A/R", 1500, cfRow)   'Collected A/R
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign


    cfRow = cfRow + 300     'SP A/R
    Set txtD = CF_Attc.Sections(3).AddTextObject("Spare Parts On-Account", 2000, cfRow)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(5000, cfRow + 280, 6600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'if SP on Acct is available display OR Series
    If SPonAcct > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(SPonAcct, "P #,##0.00"), 5000, cfRow)  'SP on acct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Qry = "SELECT Payments.ORNO From PAYMENTS Where payments.SPAREPARTS>0 AND payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL;"
        ORspOnAcct = GetOR(Qry)
        Qry = "SELECT Payments.CRNO From PAYMENTS Where payments.SPAREPARTS>0 AND payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL;"
        CRseries = GetOR(Qry)
        If ORspOnAcct <> "" Then
            ORspOnAcct = "[OR#:" & ORspOnAcct & IIf(CRseries <> "", ", CR#:" & CRseries & "]", "]")
        Else
            ORspOnAcct = IIf(CRseries <> "", "[CR#:" & CRseries & "]", "")
        End If
        cfRow = cfRow + 300
        Hx = IIf(Len(ORspOnAcct) > 90, Round((Len(ORspOnAcct) / 90 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject(ORspOnAcct, 2000, cfRow)   'SP on Acct OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If

    cfRow = cfRow + 300     'Svc On-Acct customer
    Set txtD = CF_Attc.Sections(3).AddTextObject("Service On-Account", 2000, cfRow)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(5000, cfRow + 280, 6600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'if SP on Acct is available display OR Series
    If SvcReg > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(SvcReg, "P #,##0.00"), 5000, cfRow)  'SP on acct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        
        Qry = "SELECT payments.ORNO FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND payments.INVOICENO IS NULL AND payments.CUSTOMERID > 0;"
        ORsvcOnAcct = GetOR(Qry)
        Hx = IIf(Len(ORsvcOnAcct) > 90, Round((Len(ORsvcOnAcct) / 90 + 0.49), 2), 1)
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject("[OR#: " & ORsvcOnAcct & "]", 2000, cfRow)   'SP on Acct OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If

    cfRow = cfRow + 300     'Svc FSC Acct customer
    Set txtD = CF_Attc.Sections(3).AddTextObject("Service FSC", 2000, cfRow)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(5000, cfRow + 280, 6600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'if Svc FSC Acct is available display OR Series
    If SvcFSC > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(SvcFSC, "P #,##0.00"), 5000, cfRow)  'SP on acct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        'Qry = "SELECT Payments.ORNO From PAYMENTS Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NULL AND PAYMENTS.EmpId = " & cboCashier.ItemData(cboCashier.ListIndex) & ";"
        'ORsvcFSC = GetOR(Qry)
        Hx = Round((Len(ORsvcFSC) / 90 + 0.49), 2)
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject("[OR#: " & ORsvcFSC & "]", 2000, cfRow)   'SP on Acct OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If

    cfRow = cfRow + 300     'Others On-Acct customer
    Set txtD = CF_Attc.Sections(3).AddTextObject("Others On-Account", 2000, cfRow)   'Cash Sales
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(5000, cfRow + 280, 6600, cfRow + 280) 'HLine Cash Sales
    lnObjD.LineThickness = 1
    'Display Total A/R
    TtlAR = OthrAcct + SvcFSC + SvcReg + SPonAcct
    Set txtD = CF_Attc.Sections(3).AddTextObject(Format(TtlAR, "#,##0.00"), 7000, cfRow) 'Total Collection
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine COllector
    lnObjD.LineThickness = 1
    
    'if Other Acct is available display OR Series
    If OthrAcct > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(OthrAcct, "P #,##0.00"), 5000, cfRow)  'Other on acct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Qry = "SELECT payments.ORNO FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.OTHERS > 0 AND payments.INVOICENO IS NULL AND payments.CUSTOMERID > 0;"
        ORotherAcct = GetOR(Qry)
        Qry = "SELECT payments.CRNO FROM payments WHERE payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.OTHERS > 0 AND payments.INVOICENO IS NULL AND payments.CUSTOMERID > 0;"
        ORotherAcct = GetOR(Qry)
        If ORotherAcct <> "" Then
            ORotherAcct = "[OR#:" & ORotherAcct & IIf(CRseries <> "", ", CR#:" & CRseries & "]", "]")
        Else
            ORotherAcct = IIf(CRseries <> "", "[CR#:" & CRseries & "]", "")
        End If
        Hx = IIf(Len(ORotherAcct) > 90, Round((Len(ORotherAcct) / 90 + 0.49), 2), 1)
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject(ORotherAcct, 2000, cfRow)   'Other on Acct OR Series
        txtD.Width = 7800: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If

    If CashReturn > 0 Then
        cfRow = cfRow + 300     'Svc FSC Acct customer
        Set txtD = CF_Attc.Sections(3).AddTextObject("Add: Cash Returned", 1500, cfRow)   'Cash Sales
        txtD.Width = 1300: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(5000, cfRow + 280, 6600, cfRow + 280) 'HLine Cash Sales
        lnObjD.LineThickness = 1
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(CashReturn, "P #,##0.00"), 5000, cfRow)  'Other on acct
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Qry = "SELECT cashreturn.CASHRETNO From cashreturn " & _
              "Where cashreturn.CRDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND cashreturn.PRINTCOPY > 0 AND cashreturn.REMARKS NOT LIKE '%CANCELLED%' AND cashreturn.UPDATEDBY = " & cboCashier.ItemData(cboCashier.ListIndex) & ";"
        CRseries = GetOR(Qry)
        Hx = IIf(Len(CRseries) > 0, Round((Len(CRseries) / 40 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject("[Ref#: " & CRseries & "]", 2800, cfRow)   'Other on Acct OR Series
        txtD.Width = 4000: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
    
    If Val(Replace(txtRecTOF.Text, ",", "")) > 0 Then
        cfRow = cfRow + 300
        If CashReturn > 0 Then
            Set txtD = CF_Attc.Sections(3).AddTextObject("Received Funds", 2000, cfRow) 'Add: Received funds from TOF
            txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Else
            Set txtD = CF_Attc.Sections(3).AddTextObject("Add: Received Funds", 1500, cfRow) 'Add: Received funds from TOF
            txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        End If
        Set txtD = CF_Attc.Sections(3).AddTextObject(txtRecTOF.Text, 5000, cfRow) 'Ttl Rec. TOF
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(5000, cfRow + 280, 6600, cfRow + 280) 'HLine Cash Sales
        lnObjD.LineThickness = 1
    End If
    
   'Display Total Cash Return and Received funds
    If CashReturn > 0 Or Val(Replace(txtRecTOF.Text, ",", "")) > 0 Then
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(CashReturn + Val(Replace(txtRecTOF.Text, ",", "")), "#,##0.00"), 7000, cfRow) 'Total Collection
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine COllector
        lnObjD.LineThickness = 1
    End If
    
    'Display TOTAL COLLECTION
    Set txtD = CF_Attc.Sections(3).AddTextObject(Format(TtlColltn, "#,##0.00"), 9000, cfRow) 'TOTAL COLLECTION
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine TOTALRECEIPTS
    lnObjD.LineThickness = 1
    
    'Display Less: DISBURSEMENT
    cfRow = cfRow + 500
    Set txtD = CF_Attc.Sections(3).AddTextObject("LESS:", 500, cfRow) 'LESS
    txtD.Width = 1000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set txtD = CF_Attc.Sections(3).AddTextObject("DISBURSEMENTS", 1500, cfRow) 'DISBURSEMENT
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    'Display Advances for Liquidation
    cfRow = cfRow + 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("Advances for Liquidation - Disburser", 1500, cfRow) 'Advances for Liquidation - Disburser
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Adv. Liquidation
    lnObjD.LineThickness = 1
    If AdvLiq > 0 Then 'Show Adv. for Liq.
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(AdvLiq, "#,##0.00"), 7000, cfRow) 'Adv. Liquidation
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        RowTtlDisb = cfRow
        cfRow = cfRow + 300
        Qry = "SELECT advliquidation.AFLNO From advliquidation " & _
              "Where advliquidation.AFLDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND  advliquidation.PRINTCOPY > 0 Order By advliquidation.AFLNO;"
        AdvLiqNO = GetOR(Qry)
        Hx = IIf(Len(AdvLiqNO) > 70, Round((Len(AdvLiqNO) / 70 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject("[AFL#: " & AdvLiqNO & "]", 1500, cfRow) 'AFL Series
        txtD.Width = 5500: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
    
    'Display Cash out
    cfRow = cfRow + 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("Cash Out", 1500, cfRow) 'Cash Out
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Adv. Liquidation
    lnObjD.LineThickness = 1
    If CashOutAmt > 0 Then 'Show Cash Out Amount
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(CashOutAmt, "#,##0.00"), 7000, cfRow) 'Cash Out Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Qry = "SELECT CashVoucher.VOUCHERNO From CashVoucher Where cashvoucher.CVDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND cashvoucher.REMARKS <> 'EXPENSE' AND cashvoucher.REMARKS <> 'CANCELLED';"
        CashOutNO = GetOR(Qry)
        cfRow = cfRow + 300
        Hx = IIf(Len(CashOutNO) > 70, Round((Len(CashOutNO) / 70 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject("[CV#: " & CashOutNO & "]", 1500, cfRow) 'CV Series
        txtD.Width = 5500: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
        
    'Display Sales Return
    If sRetAmt > 0 Then 'Show Sales Return Amount
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject("Sales Return", 2000, cfRow) 'Less Online Payment
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(sRetAmt, "#,##0.00"), 7000, cfRow) 'Sales REturn Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine ONLINE PAYMENT
        lnObjD.LineThickness = 1
        Qry = "SELECT sreturnsp.RETNO From sreturnsp INNER JOIN spinvoice ON (sreturnsp.INVOICE = spinvoice.SPINVNO) " & _
              "Where sreturnsp.RETDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND sreturnsp.`PROCESS` = 'REPLACE' AND spinvoice.INVTYPE = 'CASH'"
        SalRetNO = GetOR(Qry)
        Hx = IIf(Len(SalRetNO) > 70, Round((Len(SalRetNO) / 70 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject("[SR#: " & SalRetNO & "]", 1500, cfRow) 'CV Series
        txtD.Width = 5500: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
       
    'Display LESS: ONLINE PAYMENT
    If OnLineAmt > 0 Then 'Show Online Amount
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject("Online Payment", 2000, cfRow) 'Less Online Payment
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(OnLineAmt, "#,##0.00"), 7000, cfRow) 'ONLINE Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine ONLINE PAYMENT
        lnObjD.LineThickness = 1
        Qry = "SELECT PAYMENTS.ORNO From PAYMENTS " & _
                "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS LIKE '%ON-LINE%' AND PAYMENTS.INVOICENO IS NOT NULL;"
        OnlineOR = GetOR(Qry)
        Qry = "SELECT PAYMENTS.CRNO From PAYMENTS " & _
                "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS LIKE '%ON-LINE%' AND PAYMENTS.INVOICENO IS NOT NULL;"
        CRseries = GetOR(Qry)
        If OnlineOR <> "" Then
            OnlineOR = "[OR#:" & OnlineOR & IIf(CRseries <> "", ", CR#:" & CRseries, "")
        Else
            OnlineOR = IIf(CRseries <> "", "[CR#:" & CRseries, "")
        End If
        Qry = "SELECT PAYMENTS.INVOICENO From PAYMENTS " & _
                "WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS LIKE '%ON-LINE%' AND PAYMENTS.INVOICENO IS NOT NULL;"
        OnlineINV = GetOR(Qry)
        If OnlineOR <> "" Then
            OnlineOR = IIf(OnlineINV <> "", OnlineOR & ", INV#:" & OnlineINV & "]", "]")
        Else
            OnlineOR = IIf(OnlineINV <> "", "[INV#:" & OnlineINV & "]", "")
        End If
        Hx = IIf(Len(OnlineOR) > 70, Round((Len(OnlineOR) / 70 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject(OnlineOR, 1500, cfRow)  'CV Series
        txtD.Width = 5500: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
         
    'Display LESS: Credit Card PAYMENT
    If CCardAmt > 0 Then 'Show Credit CArd Amount
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject("Credit Card Payment", 2000, cfRow) 'Less Online Payment
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(CCardAmt, "#,##0.00"), 7000, cfRow) 'ONLINE Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine ONLINE PAYMENT
        lnObjD.LineThickness = 1
        Qry = "SELECT payments.INVOICENO FROM payments WHERE PAYMENTS.PAYDATE='" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND PAYMENTS.REMARKS LIKE '%CREDIT CARD%';"
        CCardOR = GetOR(Qry)
        Hx = Round((Len(CCardOR) / 70 + 0.49), 2)
        Set txtD = CF_Attc.Sections(3).AddTextObject("[INV#: " & CCardOR & "]", 1500, cfRow) 'OR Series
        'eRROR
        'txtD.Width = 5500: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        txtD.Width = 5500: txtD.Height = IIf(Hx = 0, 170, 170 * Hx): txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If
    
    'Display LESS: Inter-Branch PAYMENT
    If IBrnAmt > 0 Then 'Show Inter-Branch Amount
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject("Inter-Branch Payment", 2000, cfRow) 'Less Online Payment
        txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(IBrnAmt, "#,##0.00"), 7000, cfRow) 'ONLINE Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine ONLINE PAYMENT
        lnObjD.LineThickness = 1
        Qry = "SELECT payments.ORNO From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.REMARKS LIKE '%INTER-BRANCH%';"
        IBrnOR = GetOR(Qry)
        Qry = "SELECT payments.CRNO From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.REMARKS LIKE '%INTER-BRANCH%';"
        CRseries = GetOR(Qry)
        If IBrnOR <> "" Then
            IBrnOR = "[OR#:" & IBrnOR & IIf(CRseries <> "", ", CR#:" & CRseries, "")
        Else
            IBrnOR = IIf(CRseries <> "", "[CR#:" & CRseries, "")
        End If
        Qry = "SELECT payments.INVOICENO From payments Where payments.PAYDATE = '" & Format$(CFdate.Value, "yyyy-mm-dd") & "' AND payments.REMARKS LIKE '%INTER-BRANCH%';"
        IBrnINV = GetOR(Qry)
        If IBrnOR <> "" Then
            IBrnOR = IIf(IBrnINV <> "", IBrnOR & ", INV#:" & IBrnINV & "]", "]")
        Else
            IBrnOR = IIf(IBrnINV <> "", "[INV#:" & IBrnINV & "]", "")
        End If
        Hx = IIf(Len(IBrnOR) > 70, Round((Len(IBrnOR) / 70 + 0.49), 2), 1)
        Set txtD = CF_Attc.Sections(3).AddTextObject(IBrnOR, 1500, cfRow) 'OR Series
        txtD.Width = 5500: txtD.Height = 170 * Hx: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign: txtD.CanGrow = True
        cfRow = cfRow + (100 * Hx)
    End If

    Set txtD = CF_Attc.Sections(3).AddTextObject(Format(TtlLess, "#,##0.00"), 9000, cfRow) 'TOTAL DISBURSEMENT
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine TOTAL DISBURSEMENT
    lnObjD.LineThickness = 1
   
    'Display NET CASH ON HAND
    cfRow = cfRow + 500
    Set txtD = CF_Attc.Sections(3).AddTextObject("NET CASH ON HAND", 500, cfRow) 'NET CASH ON HAND
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    NetCash = TtlColltn - (AdvLiq + DueToMfkg + Withdrawal + TOFgmo)
    Set txtD = CF_Attc.Sections(3).AddTextObject(Format(NetCash, "#,##0.00"), 9000, cfRow) 'NET CASH ON HAND
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine NET CASH ON HAND
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = 2 'crLSDoubleLine
    'Display BREAKDOWN
    cfRow = cfRow + 500
    Set txtD = CF_Attc.Sections(3).AddTextObject("Breakdown:", 500, cfRow) 'Breakdown
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    
    '"Due PDC issued with O.R."
    cfRow = cfRow + 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("Post-Dated Checks Due", 1500, cfRow) 'pdc due
    txtD.Width = 3500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    CountPDC = CountCheck("PDC DUE")
    If CountPDC > 0 Then
        TtlPDCamt = 0
        For i = 0 To flxChk.Rows - 1
            If flxChk.TextMatrix(i, 4) = "PDC DUE" Then
                cfRow = cfRow + 300
                Set txtD = CF_Attc.Sections(3).AddTextObject("Check# " & flxChk.TextMatrix(i, 1) & " - " & flxChk.TextMatrix(i, 3) & ")", 2000, cfRow) 'Checks NO and Acct Name
                txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
                Set txtD = CF_Attc.Sections(3).AddTextObject(flxChk.TextMatrix(i, 6), 7000, cfRow) 'Check Amount
                txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
                Set lnObjD = CF_Attc.Sections(3).AddLineObject(2000, cfRow + 280, 6600, cfRow + 280) 'HLine Checks
                lnObjD.LineThickness = 1
                Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
                lnObjD.LineThickness = 1
                TtlPDCamt = TtlPDCamt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
            End If
        Next i
        'Display Check Total Amount
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(TtlPDCamt, "#,##0.00"), 9000, cfRow) 'Total PDC Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    Else
        cfRow = cfRow + 300
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(2000, cfRow + 280, 7000, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    End If

    'Display Current CHECKs
    cfRow = cfRow + 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("Current Checks", 1500, cfRow) 'Breakdown
    txtD.Width = 2500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    CountCurrent = CountCheck("CURRENT")
    If CountCurrent > 0 Then
        TtlCrntAmt = 0
        For i = 1 To flxChk.Rows - 1
            If flxChk.TextMatrix(i, 4) = "CURRENT" Then
                cfRow = cfRow + 300
                Set txtD = CF_Attc.Sections(3).AddTextObject("Check# " & flxChk.TextMatrix(i, 1) & " - " & flxChk.TextMatrix(i, 3) & ")", 2000, cfRow) 'Checks NO and Acct Name
                txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
                Set txtD = CF_Attc.Sections(3).AddTextObject(flxChk.TextMatrix(i, 6), 7000, cfRow) 'Check Amount
                txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
                Set lnObjD = CF_Attc.Sections(3).AddLineObject(2000, cfRow + 280, 6600, cfRow + 280) 'HLine Checks
                lnObjD.LineThickness = 1
                Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
                lnObjD.LineThickness = 1
                TtlCrntAmt = TtlCrntAmt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
            End If
        Next i
        'Display Check Total Amount
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(TtlCrntAmt, "#,##0.00"), 9000, cfRow) 'Total PDC Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    Else
        cfRow = cfRow + 300
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(2000, cfRow + 280, 7000, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    End If

    'Display ENCASHMENT CHECKs
    cfRow = cfRow + 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("Check Encashment:", 1500, cfRow) 'Breakdown
    txtD.Width = 2500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    CountEnCash = CountCheck("ENCASHMENT")
    If CountEnCash > 0 Then
        TtlEncAmt = 0
        For i = 0 To flxChk.Rows - 1
            If flxChk.TextMatrix(i, 4) = "ENCASHMENT" Then
                cfRow = cfRow + 300
                Set txtD = CF_Attc.Sections(3).AddTextObject("Check# " & flxChk.TextMatrix(i, 1) & " - " & flxChk.TextMatrix(i, 3) & ")", 2000, cfRow) 'Checks NO and Acct Name
                txtD.Width = 4500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
                Set txtD = CF_Attc.Sections(3).AddTextObject(flxChk.TextMatrix(i, 6), 7000, cfRow) 'Check Amount
                txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
                Set lnObjD = CF_Attc.Sections(3).AddLineObject(2000, cfRow + 280, 6600, cfRow + 280) 'HLine Checks
                lnObjD.LineThickness = 1
                Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
                lnObjD.LineThickness = 1
                TtlEncAmt = TtlEncAmt + Val(Replace(flxChk.TextMatrix(i, 6), ",", ""))
            End If
        Next i
        'Display Check Total Amount
        Set txtD = CF_Attc.Sections(3).AddTextObject(Format(TtlEncAmt, "#,##0.00"), 9000, cfRow) 'Total PDC Amount
        txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    Else
        cfRow = cfRow + 300
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(2000, cfRow + 280, 7000, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(7000, cfRow + 280, 8600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
        Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
        lnObjD.LineThickness = 1
    End If

    'Display Bills
    cfRow = cfRow + 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("Bills:", 1500, cfRow) 'Breakdown
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    Set txtD = CF_Attc.Sections(3).AddTextObject(Format(BillsTtl, "#,##0.00"), 9000, cfRow) 'Total BILLS Amount
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
    lnObjD.LineThickness = 1
    'Display Coins
    cfRow = cfRow + 300
    Set txtD = CF_Attc.Sections(3).AddTextObject("Coins:", 1500, cfRow) 'Breakdown Coins
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    Set txtD = CF_Attc.Sections(3).AddTextObject(Format(CoinsTtl, "#,##0.00"), 9000, cfRow) 'Total CoinS Amount
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
    lnObjD.LineThickness = 1
    'Display CASH ON HAND FOR DEPOSIT
    cfRow = cfRow + 400
    Set txtD = CF_Attc.Sections(3).AddTextObject("CASH ON HAND for DEPOSIT", 500, cfRow) 'CASH ON HAND FOR DEPOSIT
    txtD.Width = 4000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = True
    Set txtD = CF_Attc.Sections(3).AddTextObject(Format(Cash4Deposit, "#,##0.00"), 9000, cfRow) 'Total CoinS Amount
    txtD.Width = 1500: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crRightAlign: txtD.Font.Bold = True
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(9000, cfRow + 280, 10600, cfRow + 280) 'HLine Checks
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = 2
    
    'Encoded Remarks
    If CountPDC + CountCurrent + flxRC.Rows >= 1 Then
        cfRow = cfRow + 200
        Set txtD = CF_Attc.Sections(3).AddTextObject("Remarks: ", 500, cfRow)  'Remarks
        txtD.Width = 11000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = False
    End If
    If txtRemarks.Text <> "" Then
        cfRow = cfRow + 300
        Set txtD = CF_Attc.Sections(3).AddTextObject(txtRemarks.Text, 1000, cfRow)  'Remarks
        txtD.Width = 11000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign: txtD.Font.Bold = False
    End If
    'OR Issued to PDC
    With flxChk
        If CountPDC > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 4) = "PDC DUE" Then
                    Qry = "SELECT checkalloc.ORNO FROM checkalloc " & _
                          "WHERE  checkalloc.ORDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 7)) & " ORDER BY checkalloc.ORNO"
                    PDCor = GetOR(Qry)
                    Qry = "SELECT SUM(checkalloc.ORAmt) As SumOfORAmt From checkalloc " & _
                          "Where checkalloc.ORDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 7)) & ";"
                    ORissuedPDC = Val(GetFieldValue(Qry))
                    If ORissuedPDC > 0 Then
                        cfRow = cfRow + 250
                        Set txtD = CF_Attc.Sections(3).AddTextObject("Receipt issued to PDC Check# " & .TextMatrix(i, 2) & " [Receipt Amt: " & Format(ORissuedPDC, "#,##0.00") & " Receipt#: " & PDCor & "]", 1000, cfRow)   'OR Series PDC
                        txtD.Width = 9500: txtD.Height = 200: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign
                    End If
                End If
            Next i
        End If
        
        If CountCurrent > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 4) = "CURRENT" Then
                    Qry = "SELECT checkalloc.ORNO FROM checkalloc " & _
                          "WHERE  checkalloc.ORDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 7)) & " ORDER BY checkalloc.ORNO"
                    CrntORseries = GetOR(Qry)
                    Qry = "SELECT SUM(checkalloc.ORAmt) As SumOfORAmt From checkalloc " & _
                          "Where checkalloc.ORDATE = '" & Format(CFdate.Value, "yyyy-mm-dd") & "' AND checkalloc.CHKSPRID = " & Val(.TextMatrix(i, 7)) & ";"
                    CrntORamt = Val(GetFieldValue(Qry))
                    If CrntORamt > 0 Then
                        cfRow = cfRow + 250
                        Set txtD = CF_Attc.Sections(3).AddTextObject("Receipt issued to Current Check# " & .TextMatrix(i, 1) & " [Receipt Amt: " & Format(CrntORamt, "#,##0.00") & " Receipt#: " & CrntORseries & "]", 1000, cfRow)   'OR Series PDC
                        txtD.Width = 9500: txtD.Height = 250: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign
                    End If
                End If
            Next i
        End If
    End With
                
    'Checks Received
    If flxRC.Rows >= 1 Then
        For i = 0 To flxRC.Rows - 1
            cfRow = cfRow + 250
            Set txtD = CF_Attc.Sections(3).AddTextObject("Check received from " & flxRC.TextMatrix(i, 3) & " Check#: " & flxRC.TextMatrix(i, 1) & " Date: " & Format(flxRC.TextMatrix(i, 2)) & " Amount: " & flxRC.TextMatrix(i, 6), 1000, cfRow) 'Checks received
            txtD.Width = 9500: txtD.Height = 250: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 8: txtD.HorAlignment = crLeftAlign
        Next i
    End If
    If cfRow > 14500 Then
        cfRow = cfRow + 500
    Else
        cfRow = 15200
    End If
    Set txtD = CF_Attc.Sections(3).AddTextObject("Prepared by:", 500, cfRow)
    txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    Set txtD = CF_Attc.Sections(3).AddTextObject("Approved by:", 6500, cfRow)
    txtD.Width = 2000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crLeftAlign
    cfRow = cfRow + 400
    Set txtD = CF_Attc.Sections(3).AddTextObject(UCase$(cboCashier.Text), 2000, cfRow)
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
    Set txtD = CF_Attc.Sections(3).AddTextObject(UCase$(txtAprvby.Text), 7500, cfRow)
    txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(2000, cfRow + 280, 5000, cfRow + 280)
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = crLSDoubleLine
    Set lnObjD = CF_Attc.Sections(3).AddLineObject(7500, cfRow + 280, 10500, cfRow + 280)
    lnObjD.LineThickness = 1 ': lnObjD.LineStyle = crLSDoubleLine
    cfRow = cfRow + 300
' ******************** CODE UPDATED: 06/08/2017 *************************
    If SysBranch = "BINOS SPAREPARTS STO. TOMAS" Or SysBranch = "BINOS SPAREPARTS SAMAL" Or SysBranch = "BINOS SPAREPARTS POLOMOLOK" Or SysBranch = "BINOS SPAREPARTS TAGUM" Or SysBranch = "MOTORHUB SPAREPARTS TAGUM" Or SysBranch = "MOTORHUB SPAREPARTS MARBEL" Then
    'diri mag'edit
        Set txtD = CF_Prvw.Sections(3).AddTextObject("SPARE PARTS OIC", 2000, cfRow)    '15900)
        txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
        Set txtD = CF_Prvw.Sections(3).AddTextObject("BRANCH OIC", 7500, cfRow)    '15900)
        txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
    Else
        Set txtD = CF_Attc.Sections(3).AddTextObject("MANAGER", 2000, cfRow)
        txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
        Set txtD = CF_Attc.Sections(3).AddTextObject("AOM", 7500, cfRow)
        txtD.Width = 3000: txtD.Height = 300: txtD.Font.Name = "TAHOMA": txtD.Font.Size = 10: txtD.HorAlignment = crHorCenterAlign
    End If
' **************** END CODE UPDATED: 06/08/2017 *************************
    
    'Section 1: Report Footer
    CF_Attc.Sections(4).Suppress = True
    'Section 2: Page Header
    CF_Attc.Sections(5).Suppress = True
    
    
    'Display Report in Viewer
    CRViewer2.ReportSource = CF_Attc
    CRViewer2.DisplayGroupTree = False
    CRViewer2.ViewReport
    CRViewer2.Zoom 100

    SSTab1.TabCaption(2) = "Attachment"
    SSTab1.TabEnabled(2) = True

End Sub


Private Function CountCheck(chkType As String) As Integer
    Dim i As Integer, myCount As Integer
    
    myCount = 0
    With flxChk
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 4) = chkType Then myCount = myCount + 1
        Next i
    End With
    CountCheck = myCount
End Function


Private Sub txtCVdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtCVDATE.Tag = "skiplostfocus"
        txtCVDATE.Visible = False
        flxTOF.SetFocus
        flxTOF.Col = 1
    End If
End Sub

Private Sub txtCVDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890/"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxTOF
            If .TextMatrix(ActRow, 3) <> "" And .TextMatrix(ActRow, 2) <> "" And .TextMatrix(ActRow, 4) <> "" And ActRow = .Rows - 1 Then
                .AddItem ""
                .Col = 1
                .SetFocus
            Else
                flxTOF.Col = 2
                flxTOF.Row = ActRow
                flxTOF.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtCVDate_LostFocus()
    If txtCVDATE.Tag = "skiplostfocus" Then
        txtCVDATE.Tag = ""
        Exit Sub
    End If
    flxTOF.TextMatrix(ActRow, ActCol) = txtCVDATE.Text
    txtCVDATE.Visible = False
End Sub

Private Sub txtCVNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtCVNum.Tag = "skiplostfocus"
        txtCVNum.Visible = False
        flxTOF.SetFocus
        flxTOF.Col = ActCol
    End If
End Sub

Private Sub txtCVNum_KeyPress(KeyAscii As Integer)
    Dim strValid As String, x As Integer
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxTOF
            x = IIf(ActCol = 2, 4, 2)
            If .TextMatrix(ActRow, x) <> "" And .TextMatrix(ActRow, 3) <> "" And ActRow = .Rows - 1 Then
                .AddItem ""
                .Col = 1
                .SetFocus
            Else
                flxTOF.Col = ActCol
                flxTOF.Row = ActRow
                flxTOF.SetFocus
            End If
        End With
    End If

End Sub

Private Sub txtCVNum_LostFocus()
    Dim i As Integer, myTtl As Currency
    
    If txtCVNum.Tag = "skiplostfocus" Then
        txtCVNum.Tag = ""
        Exit Sub
    End If
    flxTOF.TextMatrix(ActRow, ActCol) = IIf(ActCol = 4, Format(txtCVNum.Text, "#,##0.00"), txtCVNum.Text)
    txtCVNum.Visible = False
    If ActCol = 4 Then
        myTtl = 0
        For i = 1 To flxTOF.Rows - 1
            myTtl = myTtl + Val(Replace(flxTOF.TextMatrix(i, 4), ",", ""))
        Next i
        txtTtlTOF.Text = Format(myTtl, "#,##0.00")
    End If
    
End Sub

Private Sub txtRecFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtRecFrom.Tag = "skiplostfocus"
        txtRecFrom.Visible = False
        flxTOF.SetFocus
        flxTOF.Col = ActCol
    End If
End Sub

Private Sub txtRecFrom_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        With flxTOF
            If .TextMatrix(ActRow, 1) <> "" And .TextMatrix(ActRow, 2) <> "" And .TextMatrix(ActRow, 3) <> "" And ActRow = .Rows - 1 Then
                .AddItem ""
                .Col = 1
                .SetFocus
            Else
                flxTOF.Col = ActCol
                flxTOF.Row = ActRow
                flxTOF.SetFocus
            End If
        End With
    End If
End Sub

Private Sub txtRecFrom_LostFocus()
    If txtRecFrom.Tag = "skiplostfocus" Then
        txtRecFrom.Tag = ""
        Exit Sub
    End If
    flxTOF.TextMatrix(ActRow, ActCol) = txtRecFrom.Text
    txtRecFrom.Visible = False
End Sub

Private Sub txtUndpColl_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub
