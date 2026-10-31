VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SPSorder 
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   -27120
   ClientWidth     =   12870
   ControlBox      =   0   'False
   Icon            =   "SPSorderWSD.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   12870
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabFinal 
      Height          =   1095
      Left            =   5160
      TabIndex        =   141
      Tag             =   "xyf"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "VERIFICATION CODE"
      TabPicture(0)   =   "SPSorderWSD.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCan"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdVerify"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox txtCode 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   144
         Tag             =   "xyf"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "OK"
         Height          =   255
         Left            =   240
         TabIndex        =   143
         Tag             =   "xyf"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCan 
         Caption         =   "CANCEL"
         Height          =   255
         Left            =   1440
         TabIndex        =   142
         Tag             =   "xyf"
         Top             =   720
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab tabKA 
      Height          =   4455
      Left            =   13320
      TabIndex        =   134
      Top             =   7200
      Visible         =   0   'False
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPSorderWSD.frx":08E6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdKACancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdKAOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboKA"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.ComboBox cboKA 
         Height          =   2910
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   137
         Top             =   480
         Width           =   3615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "CASH"
         Height          =   255
         Left            =   2040
         TabIndex        =   136
         Top             =   3480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ON-ACCOUNT"
         Height          =   255
         Left            =   360
         TabIndex        =   135
         Top             =   3510
         Value           =   -1  'True
         Width           =   1695
      End
      Begin lvButton.lvButtons_H cmdKAOK 
         Height          =   375
         Left            =   240
         TabIndex        =   138
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "SPSorderWSD.frx":0902
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdKACancel 
         Height          =   375
         Left            =   2040
         TabIndex        =   139
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "SPSorderWSD.frx":0E9C
         cBack           =   -2147483633
      End
      Begin VB.Label Label19 
         Caption         =   "Select Key Accounts:"
         Height          =   255
         Left            =   480
         TabIndex        =   140
         Top             =   120
         Width           =   2055
      End
   End
   Begin lvButton.lvButtons_H cmdRemoveWS 
      Height          =   375
      Left            =   6240
      TabIndex        =   89
      Tag             =   "XYF"
      ToolTipText     =   "Remove Customer"
      Top             =   6180
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
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
      Image           =   "SPSorderWSD.frx":1436
      cBack           =   -2147483633
   End
   Begin TabDlg.SSTab tabPromoCard 
      Height          =   2775
      Left            =   4560
      TabIndex        =   113
      Top             =   2760
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "EARN POINTS"
      TabPicture(0)   =   "SPSorderWSD.frx":19D0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdPCCancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPCOK"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPCCustomer"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPCCNo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPCCPoints"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboPCPointsType"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "REDEEM POINTS"
      TabPicture(1)   =   "SPSorderWSD.frx":19EC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(4)=   "cmdRPCancel"
      Tab(1).Control(5)=   "cmdRPOK"
      Tab(1).Control(6)=   "txtRPCardPoints"
      Tab(1).Control(7)=   "txtRPCardNo"
      Tab(1).Control(8)=   "txtRPCustomer"
      Tab(1).Control(9)=   "cboRPPointsType"
      Tab(1).ControlCount=   10
      Begin VB.ComboBox cboPCPointsType 
         Height          =   315
         Left            =   1800
         TabIndex        =   133
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox cboRPPointsType 
         Height          =   315
         Left            =   -73200
         TabIndex        =   131
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtRPCustomer 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73200
         TabIndex        =   124
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtRPCardNo 
         Height          =   375
         Left            =   -73200
         TabIndex        =   123
         Top             =   1025
         Width           =   2775
      End
      Begin VB.TextBox txtRPCardPoints 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73200
         TabIndex        =   122
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtPCCPoints 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   119
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtPCCNo 
         Height          =   375
         Left            =   1800
         TabIndex        =   118
         Top             =   1020
         Width           =   2775
      End
      Begin VB.TextBox txtPCCustomer 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   117
         Top             =   600
         Width           =   4335
      End
      Begin lvButton.lvButtons_H cmdPCOK 
         Height          =   375
         Left            =   1680
         TabIndex        =   120
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "SPSorderWSD.frx":1A08
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPCCancel 
         Height          =   375
         Left            =   3480
         TabIndex        =   121
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "SPSorderWSD.frx":1FA2
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRPOK 
         Height          =   375
         Left            =   -73320
         TabIndex        =   125
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "SPSorderWSD.frx":253C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRPCancel 
         Height          =   375
         Left            =   -71520
         TabIndex        =   126
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "SPSorderWSD.frx":2AD6
         cBack           =   -2147483633
      End
      Begin VB.Label Label18 
         Caption         =   "POINTS TYPE:"
         Height          =   255
         Left            =   240
         TabIndex        =   132
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "POINTS TYPE:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   130
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "CUSTOMER NAME:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   129
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "CARD NUMBER:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   128
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "CARD POINTS:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   127
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "CARD POINTS:"
         Height          =   255
         Left            =   240
         TabIndex        =   116
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "CARD NUMBER:"
         Height          =   255
         Left            =   240
         TabIndex        =   115
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "CUSTOMER NAME:"
         Height          =   255
         Left            =   240
         TabIndex        =   114
         Top             =   660
         Width           =   1575
      End
   End
   Begin VB.TextBox txtYClub 
      Height          =   615
      Left            =   13320
      TabIndex        =   100
      Text            =   "10"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkClub 
      Caption         =   "Yamaha Club member"
      Height          =   255
      Left            =   360
      TabIndex        =   99
      Tag             =   "xyf"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtTotPen 
      Height          =   285
      Left            =   2400
      TabIndex        =   94
      Top             =   10620
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H000000FF&
      Height          =   1395
      Left            =   240
      TabIndex        =   91
      Tag             =   "XYF"
      Top             =   7110
      Width           =   12495
      Begin MSFlexGridLib.MSFlexGrid flxOrder 
         Height          =   345
         Left            =   60
         TabIndex        =   92
         Tag             =   "XYF"
         Top             =   930
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   609
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         ForeColor       =   255
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         ScrollBars      =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxGridRef 
         Height          =   495
         Left            =   105
         TabIndex        =   95
         Tag             =   "XYF"
         Top             =   120
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   873
         _Version        =   393216
         Rows            =   8
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   3
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         GridColorFixed  =   16777215
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   1
         BorderStyle     =   0
         FormatString    =   $"SPSorderWSD.frx":3070
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PENDING ORDER"
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
         Left            =   60
         TabIndex        =   96
         Tag             =   "XYF"
         Top             =   630
         Width           =   1215
      End
   End
   Begin VB.TextBox txtVat 
      Height          =   285
      Left            =   2760
      TabIndex        =   88
      Top             =   10230
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPec 
      Height          =   615
      Left            =   2040
      TabIndex        =   87
      Top             =   10350
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CheckBox chkWholesaleCash 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1440
      TabIndex        =   86
      Top             =   10950
      Visible         =   0   'False
      Width           =   1335
   End
   Begin TabDlg.SSTab tabWS 
      Height          =   4455
      Left            =   13320
      TabIndex        =   79
      Top             =   1680
      Visible         =   0   'False
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPSorderWSD.frx":3147
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdWSCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdWSOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboWS"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptCash"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OptOnAcct"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.OptionButton OptOnAcct 
         Caption         =   "ON-ACCOUNT"
         Height          =   255
         Left            =   360
         TabIndex        =   85
         Top             =   3510
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptCash 
         Caption         =   "CASH"
         Height          =   255
         Left            =   2040
         TabIndex        =   84
         Top             =   3480
         Width           =   975
      End
      Begin VB.ComboBox cboWS 
         Height          =   2910
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   80
         Top             =   480
         Width           =   3615
      End
      Begin lvButton.lvButtons_H cmdWSOK 
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "SPSorderWSD.frx":3163
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdWSCancel 
         Height          =   375
         Left            =   2040
         TabIndex        =   82
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "SPSorderWSD.frx":36FD
         cBack           =   -2147483633
      End
      Begin VB.Label Label6 
         Caption         =   "Select Wholesale:"
         Height          =   255
         Left            =   480
         TabIndex        =   83
         Top             =   120
         Width           =   2055
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxWDisc 
      Height          =   735
      Left            =   8760
      TabIndex        =   71
      Top             =   10950
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtCusType 
      Height          =   285
      Left            =   360
      TabIndex        =   70
      Top             =   10350
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4560
      Top             =   1440
   End
   Begin VB.TextBox txtCPOid 
      Height          =   285
      Left            =   5280
      TabIndex        =   69
      Top             =   10590
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   67
      Top             =   9630
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CheckBox chkDP 
      Caption         =   "DP"
      Height          =   255
      Left            =   360
      TabIndex        =   66
      Top             =   10830
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstDelete 
      Height          =   645
      ItemData        =   "SPSorderWSD.frx":3C97
      Left            =   8760
      List            =   "SPSorderWSD.frx":3C99
      TabIndex        =   65
      Top             =   10230
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   1800
      TabIndex        =   64
      Top             =   10230
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmpID 
      Height          =   285
      Left            =   10920
      TabIndex        =   62
      Top             =   9390
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fraAdmin 
      Height          =   735
      Left            =   7080
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cboOrderNo 
         Height          =   315
         Left            =   4200
         TabIndex        =   61
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboPartsMan 
         Height          =   315
         Left            =   960
         TabIndex        =   59
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Order No."
         Height          =   255
         Left            =   3360
         TabIndex        =   60
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Partsman:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.TextBox txtOSBal 
      Height          =   285
      Left            =   6480
      TabIndex        =   56
      Top             =   10950
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   5040
      TabIndex        =   55
      Top             =   10800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPONO 
      Height          =   285
      Left            =   7200
      TabIndex        =   54
      Top             =   9390
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtBillAmt 
      Height          =   285
      Left            =   4200
      TabIndex        =   53
      Top             =   10950
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayStart 
      Height          =   285
      Left            =   3360
      TabIndex        =   52
      Top             =   10950
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtNPays 
      Height          =   285
      Left            =   2520
      TabIndex        =   51
      Top             =   10950
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayScheme 
      Height          =   285
      Left            =   1560
      TabIndex        =   50
      Top             =   10950
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTerms 
      Height          =   285
      Left            =   480
      TabIndex        =   49
      Top             =   10950
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCustName 
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
      TabIndex        =   47
      Tag             =   "XYF"
      Top             =   6210
      Width           =   4335
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   360
      TabIndex        =   44
      Top             =   10230
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
      TabIndex        =   43
      Top             =   9870
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
      TabIndex        =   42
      Text            =   "No Discount"
      Top             =   9750
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtGrandTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   6465
      Width           =   1620
   End
   Begin VB.TextBox txtLessDiscAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   38
      Tag             =   "XYF"
      Top             =   6180
      Width           =   1620
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   37
      Tag             =   "XYF"
      Top             =   5895
      Width           =   1620
   End
   Begin VB.TextBox txtPctDisc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   34
      Tag             =   "XYF"
      Text            =   "0%"
      Top             =   6210
      Width           =   495
   End
   Begin VB.TextBox txtTtlDisc 
      Height          =   285
      Left            =   9840
      TabIndex        =   32
      Top             =   9390
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8760
      TabIndex        =   31
      Top             =   9390
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   8040
      TabIndex        =   30
      Top             =   9390
      Visible         =   0   'False
      Width           =   615
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
      Left            =   1920
      TabIndex        =   29
      Tag             =   "XYF"
      Top             =   5895
      Width           =   1695
   End
   Begin VB.TextBox txtInvNo 
      Height          =   285
      Left            =   5760
      TabIndex        =   27
      Top             =   9390
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid flxPND 
      Height          =   600
      Left            =   5400
      TabIndex        =   26
      Top             =   150
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1058
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   16777215
      BackColorBkg    =   12648447
      GridColor       =   16777215
      GridColorFixed  =   16777215
      ScrollBars      =   1
      BorderStyle     =   0
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   315
      Left            =   9720
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   1440
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
      Image           =   "SPSorderWSD.frx":3C9B
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   10800
      TabIndex        =   24
      Tag             =   "XYF"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtSPSorderID 
      Height          =   285
      Left            =   480
      TabIndex        =   23
      Top             =   9390
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   240
      TabIndex        =   21
      Tag             =   "XYF"
      Top             =   8520
      Width           =   12495
      Begin VB.TextBox txtFrnt 
         BackColor       =   &H80000003&
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   240
         Width           =   4335
      End
      Begin lvButton.lvButtons_H cmdCreditCard 
         Height          =   495
         Left            =   9880
         TabIndex        =   102
         Tag             =   "XYF"
         ToolTipText     =   "Cashier"
         Top             =   165
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":4235
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdWholesale 
         Height          =   495
         Left            =   9250
         TabIndex        =   78
         Tag             =   "XYF"
         ToolTipText     =   "[F10] Insert Wholesale"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":4585
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRemoveCC 
         Height          =   495
         Left            =   9885
         TabIndex        =   109
         Tag             =   "XYF"
         ToolTipText     =   "Cancel Order"
         Top             =   165
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":491F
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPO 
         Height          =   495
         Left            =   10520
         TabIndex        =   77
         Tag             =   "XYF"
         ToolTipText     =   "[F11] Insert On Account"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":57F9
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdReq 
         Height          =   495
         Left            =   11160
         TabIndex        =   76
         Tag             =   "XYF"
         ToolTipText     =   "Insert Requisition"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":5D93
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdAdvFind 
         Height          =   495
         Left            =   11760
         TabIndex        =   72
         Tag             =   "XYF"
         ToolTipText     =   "[F12] Advance Search"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":632D
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   7990
         TabIndex        =   75
         Tag             =   "XYF"
         ToolTipText     =   "Cancel Order"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":68C7
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCashier 
         Height          =   495
         Left            =   8640
         TabIndex        =   73
         Tag             =   "XYF"
         ToolTipText     =   "[F9] KEY ACCOUNTS"
         Top             =   165
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":77A1
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdWareHouse 
         Height          =   495
         Left            =   7320
         TabIndex        =   74
         Tag             =   "XYF"
         ToolTipText     =   "Warehouse"
         Top             =   165
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "SPSorderWSD.frx":7E77
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FRONTLINER:"
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
         Left            =   120
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   292
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   4920
      TabIndex        =   20
      Top             =   9390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   9630
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   10110
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   9630
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   10110
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Top             =   9390
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1815
      Left            =   3000
      TabIndex        =   16
      Top             =   9480
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
      TabIndex        =   15
      Top             =   9750
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
   Begin VB.TextBox txtReqID 
      Height          =   285
      Left            =   6240
      TabIndex        =   14
      Top             =   9390
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Tag             =   "XY"
      Top             =   10920
      Width           =   135
   End
   Begin VB.Timer tmrQtyTip 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   9750
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "Qty. On Hand: 0"
      Top             =   9870
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtFlx 
      Height          =   315
      Left            =   3720
      TabIndex        =   11
      Top             =   9390
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   10590
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   240
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   720
      Width           =   12615
      Begin VB.TextBox txtOrderno 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   180
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   315
         Left            =   9240
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER NO."
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
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   210
         Width           =   975
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
         Left            =   7680
         TabIndex        =   5
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
            Picture         =   "SPSorderWSD.frx":8411
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderWSD.frx":92EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderWSD.frx":A1C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderWSD.frx":B09F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPSorderWSD.frx":BF79
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
      Image           =   "SPSorderWSD.frx":CE53
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   2
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
      Image           =   "SPSorderWSD.frx":D3ED
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Save Order"
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
      Image           =   "SPSorderWSD.frx":D987
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Cancel Order"
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
      Image           =   "SPSorderWSD.frx":DF21
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
      Image           =   "SPSorderWSD.frx":E4BB
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxBO 
      Height          =   1095
      Left            =   7560
      TabIndex        =   68
      Top             =   10830
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
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   1920
      TabIndex        =   98
      Tag             =   "XYF"
      Top             =   6525
      Width           =   4335
   End
   Begin TabDlg.SSTab tabCommandCC 
      Height          =   1215
      Left            =   5160
      TabIndex        =   110
      Top             =   3720
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2143
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "SELECT AN ACTION"
      TabPicture(0)   =   "SPSorderWSD.frx":EA55
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSubPromoCard"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSubCreditCard"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin lvButton.lvButtons_H cmdSubCreditCard 
         Height          =   495
         Left            =   360
         TabIndex        =   111
         Tag             =   "XYF"
         ToolTipText     =   "Cashier"
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Caption         =   "CREADIT CARD"
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
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "SPSorderWSD.frx":EA71
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdSubPromoCard 
         Height          =   495
         Left            =   2520
         TabIndex        =   112
         Tag             =   "XYF"
         ToolTipText     =   "Cashier"
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Caption         =   "PROMO CARD"
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
         cBhover         =   12648447
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "SPSorderWSD.frx":EDC1
         cBack           =   -2147483633
      End
   End
   Begin TabDlg.SSTab ssCreditTab 
      Height          =   1695
      Left            =   5400
      TabIndex        =   103
      Top             =   3720
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2990
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "BANK CHARGES PERCENTAGE"
      TabPicture(0)   =   "SPSorderWSD.frx":F111
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCCCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCCOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtBankCharge"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtBankCharge 
         Height          =   285
         Left            =   2640
         TabIndex        =   105
         Top             =   600
         Width           =   975
      End
      Begin lvButton.lvButtons_H cmdCCOk 
         Height          =   375
         Left            =   480
         TabIndex        =   106
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "SPSorderWSD.frx":F12D
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCCCancel 
         Height          =   375
         Left            =   2160
         TabIndex        =   107
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "SPSorderWSD.frx":F6C7
         cBack           =   -2147483633
      End
      Begin VB.Label Label10 
         Caption         =   "%"
         Height          =   255
         Left            =   3720
         TabIndex        =   108
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Please Enter Additional Charges:"
         Height          =   255
         Left            =   240
         TabIndex        =   104
         Top             =   600
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   4065
      Left            =   240
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   1800
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7170
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   $"SPSorderWSD.frx":FC61
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
   Begin VB.Label lblToolTip 
      Height          =   255
      Left            =   10080
      TabIndex        =   101
      Tag             =   "xyf"
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Customer Address"
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
      Left            =   360
      TabIndex        =   97
      Tag             =   "XYF"
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lblSPNumber 
      Height          =   315
      Left            =   210
      TabIndex        =   93
      Tag             =   "xyf"
      Top             =   6900
      Width           =   1695
   End
   Begin VB.Label lblTerms 
      Height          =   255
      Left            =   1920
      TabIndex        =   90
      Tag             =   "XYF"
      Top             =   6930
      Width           =   4815
   End
   Begin VB.Label lblRemarks 
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
      Left            =   3840
      TabIndex        =   48
      Tag             =   "XYF"
      Top             =   5970
      Width           =   2325
   End
   Begin VB.Label lblCustomer 
      Caption         =   "Customer"
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
      Left            =   360
      TabIndex        =   46
      Tag             =   "XYF"
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblRef 
      Caption         =   "Reference No."
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
      Left            =   360
      TabIndex        =   45
      Tag             =   "XYF"
      Top             =   5925
      Width           =   1935
   End
   Begin VB.Label lblGrandTtl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "  GRAND TOTAL:  "
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
      Left            =   9750
      TabIndex        =   36
      Tag             =   "XYF"
      Top             =   6495
      Width           =   1335
   End
   Begin VB.Label lblDiscount 
      AutoSize        =   -1  'True
      Caption         =   "DISCOUNT:  "
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
      Left            =   10080
      TabIndex        =   35
      Tag             =   "XYF"
      Top             =   6210
      Width           =   900
   End
   Begin VB.Label lblLess 
      AutoSize        =   -1  'True
      Caption         =   "Less:  "
      Height          =   195
      Left            =   9000
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   6225
      Width           =   465
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
      TabIndex        =   28
      Top             =   6090
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   10800
      Top             =   9750
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13110
      m_MDI_Height    =   9645
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
      Left            =   10320
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   5940
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
      TabIndex        =   7
      Top             =   1545
      Width           =   2415
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
      TabIndex        =   8
      Top             =   1575
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   7335
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
Attribute VB_Name = "SPSorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsParts As New ADODB.Recordset

Dim PrtSrch As Boolean, ActSrch As Boolean, PrtBln As Boolean, ProperExit As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim tmrCount As Integer, BtnIndx As Integer
Public AddNew As Boolean, ApplyWholeSalePriceInRetail As Boolean, ApplyKeyAcctsPriceInRetail As Boolean

Dim PmanSrch As Boolean
Public JustScroll As Boolean, StatSrch As Boolean
Dim fRow As Integer, fCol As Integer, WSdiscountID As Integer, KAdiscountID As Integer
Public myCardNo As String, myCardPointsType As String
Public myCardCustomerID As Long
Public EMPID As Integer


Private Sub ClearEntries()
    InsertPO.isKeyAcc = False
    txtOrderno.Text = "<Auto>"
    'txtDate.Text = Format(Date, "##/##/####")
    lblRemarks.Caption = ""
    lblRef.Caption = "Reference No."
    lblCustomer.Caption = "Customer"
    txtTotal.Text = ""
    txtSPSorderID.Text = ""
    txtInvNo.Text = ""
    txtReqID.Text = ""
    txtPOno.Text = ""
    txtRefNo.Text = ""
    txtCustID.Text = ""
    txtTtlDisc.Text = ""
    txtTerms.Text = ""
    txtPayScheme.Text = ""
    txtNPays.Text = ""
    txtPayStart.Text = ""
    txtBillAmt.Text = ""
    txtOSBal.Text = ""
    txtAddress.Text = ""
    txtCustName.Text = ""
    cboAddress.Text = ""
    txtGrandTotal.Text = ""
    txtLessDiscAmt.Text = ""
    flxOrd.Rows = 1
    flxOrd.Rows = 2
    txtPrintCopy = ""
    txtCusType.Text = ""
    lblTerms.Caption = ""
    cmdRemoveWS.Visible = False
    chkDP.Value = 0
    txtPctDisc.Text = ""
    
    myCardNo = ""
    myCardCustomerID = 0
    myCardPointsType = ""
End Sub
                                                                
Private Sub LockData(bln As Boolean)
    If UserType = "Administrator" Then
        fraAdmin.Visible = bln
    End If
    JustScroll = bln
    cboAddress.Locked = bln
    txtCustName.Locked = bln
End Sub
   
Private Sub cboAddress_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    With cboAddress
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'txtSrcID.Tag = ""
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            'txtSrcID.Tag = "invalid"
        End If
    End With

End Sub

Private Sub cboAddress_Click()
    Call cboAddress_Change
End Sub

Private Sub cboAddress_GotFocus()
   Dim Tmp

   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboAddress.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub cboKEYACCTS_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboWS
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'txtCustID.Text = .ItemData(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
        End If
    End With
End Sub


Private Sub cboKA_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboKA
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'txtCustID.Text = .ItemData(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
        End If
    End With
End Sub

Private Sub cboOrderNo_Change()
    Call flxPND_Click
End Sub

Private Sub cboOrderNo_Click()
    Call flxPND_Click
End Sub

Private Sub cboOrderNo_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
    If KeyCode = vbKeyF9 And Not JustScroll Then Call cmdCashier_Click 'Key Accounts
End Sub

Private Sub cboPartsMan_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboPartsMan
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            txtEmpID.Text = cboPartsMan.ItemData(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtEmpID.Text = ""
        End If
    End With
    cboOrderNo.Clear
    cboOrderNo.Text = ""
    Call ClearEntries
End Sub

Private Sub cboPartsMan_Click()
    Call cboPartsMan_Change
End Sub

Private Sub cboPartsMan_GotFocus()
    Dim Tmp
    
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cboPartsMan.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboPartsMan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
End Sub

Private Sub cboPartsMan_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
    If KeyCode = vbKeyF9 And Not JustScroll Then Call cmdCashier_Click 'Key Accounts
End Sub

Private Sub cboPartsMan_LostFocus()
    Dim rsPmanID As New ADODB.Recordset
    Dim Qry As String
    
    If txtEmpID.Text <> "" Then
        'Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & txtEmpID.Text & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#)) ORDER BY ORDERNO;"
        'mySQL
         Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & Val(txtEmpID.Text) & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(txtDate.Text, "yyyy-mm-dd") & "')) ORDER BY ORDERNO;"
       
        With rsPmanID
            cboOrderNo.Clear
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    cboOrderNo.AddItem !ORDERNO
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
End Sub

Private Sub cboPCPointsType_Click()
If txtPCCNo.Text <> "" Then
    Call txtPCCNo_LostFocus
End If
End Sub

Private Sub cboRPPointsType_Click()
If txtRPCardNo.Text <> "" Then
    Call txtRPCardNo_LostFocus
End If
End Sub

Private Sub cboWS_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboWS
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'txtCustID.Text = .ItemData(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
        End If
    End With
End Sub

Private Sub cboWS_Click()
    Call cboWS_Change
End Sub

Private Sub cboWS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
End Sub

Private Sub cmdAdvFind_Click()
    SPfind.MySource = "SPSORDER"
    SPfind.Show vbModal
    'MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Order"
End Sub

Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdCan_Click()
    tabFinal.Visible = False
    tabFinal.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Dim Ans As Integer, eMsg As String
    
    On Error GoTo LocalErr
    If txtOrderno.Text = "" Or Not IsNumeric(txtOrderno.Text) Then Exit Sub
    Ans = MsgBox("Do you want to Cancel Order No. " & txtOrderno.Text & "?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    eMsg = "Updating Sales Order for Cancellation"
    'LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                   "WHERE (((SPSALESORDER.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                   "WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    eMsg = "Updating Sales Order Details for Cancellation"
    'LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.SPSORDERID = 0 WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.SPSORDERID = 0 WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    MsgBox "Order successfully Cancelled!", vbInformation + vbOKOnly, "Cancel"
    Call cboPartsMan_Change
    Call ClearEntries
    If UserType <> "Administrator" Then Call DisplayPendingOrders
    Exit Sub
LocalErr:
    MsgBox "Error Updating table SP Sales Order during Cancel" & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdCashier_Click() 'for Key Accounts
    
''    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Then MyErr = MyErr & vbCrLf & "Please Add Item First!" End If
    
'    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER WHERE CUSTOMER.CTYPE='KEY ACCOUNTS' ORDER BY CUSTOMER.CUSTOMERNAME;"
'
'    cboKA.Clear
'    With rsKA
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                cboKA.AddItem !CustomerName
'                cboKA.ItemData(cboKA.ListCount - 1) = !CUSTOMERID
'                .MoveNext
'            Loop
'        End If
'        .Close
'    End With
'    flxOrd.Enabled = False
'    Frame2.Enabled = False
'    cmdTrans(3).Enabled = False
'    cmdTrans(4).Enabled = False
'    txtRefNo.Locked = True
'    With tabKA
'        .Caption = ""
'        .Top = (Me.Height - .Height) / 2
'        .Left = (Me.Width - .Width) / 2
'        .Visible = True
'        .SetFocus
'    End With
    With tabFinal
        .Top = (Me.Height - .Height) / 2
        .Left = (Me.Width - .Width) / 2
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub cmdCCCancel_Click()
    flxOrd.Enabled = True
    Frame2.Enabled = True
    ssCreditTab.Visible = False
    ssCreditTab.Top = 13000
    cmdTrans(3).Enabled = True
    cmdTrans(4).Enabled = True
End Sub

Private Sub cmdCCOk_Click()
Dim ttlBankCharge As Currency
Dim ttlPercDecimal As Currency
Dim ttlTotalPrice As Currency
If Val(txtBankCharge.Text) > 100 Then
    MsgBox "You Cannot Exceed to 100% for Bank Charge", vbInformation + vbOKOnly, "Validation"
    Exit Sub
End If
ttlPercDecimal = ((Val(txtBankCharge.Text)) / 100)
For i = 1 To flxOrd.Rows - 1
    If flxOrd.TextMatrix(i, 4) <> "" Then
        'flxOrd.TextMatrix(i, 8) = Val(Replace(flxOrd.TextMatrix(i, 7), ",", "")) * Val(flxOrd.TextMatrix(i, 4))
        ttlBankCharge = (Val(Replace(flxOrd.TextMatrix(i, 7), ",", "")) * ttlPercDecimal)
        ttlTotalPrice = (Val(Replace(flxOrd.TextMatrix(i, 7), ",", "")) + ttlBankCharge)
        flxOrd.TextMatrix(i, 7) = Format(Round(Format(roundUp(ttlTotalPrice), "#,###,##0.00")), "#,###,##0.00")
    End If
    ttlTotalPrice = 0
    ttlBankCharge = 0
Next i
ttlPercDecimal = 0
flxOrd.Enabled = True
Frame2.Enabled = True
cmdRemoveCC.Visible = True
cmdCreditCard.Visible = False
ssCreditTab.Visible = False
ssCreditTab.Top = 13000
cmdTrans(3).Enabled = True
cmdTrans(4).Enabled = True
Call ComputeTotalwithBankCharge
End Sub

Public Function roundUp(dblValue As Currency) As Currency
On Error GoTo er
Dim myDec As Long
myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
    roundUp = CDbl(Left(CStr(dblValue), myDec)) + 1
Else
    roundUp = dblValue
End If

Exit Function
er:
    MsgBox Err.Description, vbInformation, "Round Up"
End Function

Public Sub ComputeTotalwithBankCharge()
Dim GrndTtlPrice As Currency
    With flxOrd
        GrndTtlPrice = 0
        TtlDiscAmt = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 4) <> "" Then
                .TextMatrix(i, 8) = Format(Val(.TextMatrix(i, 2)) * Replace(.TextMatrix(i, 7), ",", ""), "#,##0.00")
                GrndTtlPrice = GrndTtlPrice + Val(Replace(.TextMatrix(i, 8), ",", ""))
            End If
        Next i
        txtTotal.Text = Format$(GrndTtlPrice, "#,##0.00")
        txtGrandTotal.Text = Format$(GrndTtlPrice - Val(Replace(txtLessDiscAmt.Text, ",", "")), "#,##0.00")
    End With
End Sub
Private Sub cmdCreditCard_Click()
flxOrd.Enabled = False
Frame2.Enabled = False
cmdTrans(3).Enabled = False
cmdTrans(4).Enabled = False
txtBankCharge.Text = "0"
With tabCommandCC
    .Top = (Me.Height - .Height) / 2
    .Left = (Me.Width - .Width) / 2
    .Visible = True
    .SetFocus
End With
'MsgBox "Not Yet Available", vbInformation + vbOKOnly, "Validation" cmdSubCreditCard tabCommandCC
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSI
        For i = fRow To .Rows - 1
            For j = fCol To 9
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 9:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 9, i + 1, i)
                    fCol = IIf(j = 9, 1, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
End Sub

Private Sub cmdKACancel_Click()
    flxOrd.Enabled = True
    Frame2.Enabled = True
    tabKA.Visible = False
    tabKA.Top = 13000
    cmdTrans(3).Enabled = True
    cmdTrans(4).Enabled = True
End Sub
Private Sub cmdKAOK_Click() 'for Key Accounts
    Dim rsDisc As New ADODB.Recordset
    Dim Qry As String, Ans As Integer, i As Integer, PctDisc As Currency
    Dim NewPrc As Currency, WDisc As Currency, PrevTtl As Currency
    Dim myDscnt As Currency, myDscntLess As Currency
    Dim TtlOil As Integer, TtlTire As Integer, NonOilTire As Integer
    
    If cboKA.ListIndex = -1 Then
        MsgBox "Please select a customer from the list!", vbExclamation + vbOKOnly, "KEY ACCOUNTS"
        Exit Sub
    End If
    
'    'Validate user if allowed to Process
'     Qry = "SELECT SETTINGS.REMARKS From SETTINGS " & _
'              "WHERE SETTINGS.DESCRIPTION = 'BM DISCOUNT' AND SETTINGS.VALUE = '8/01/2016';"
'     KAdiscountID = Val(GetFieldValue(Qry))

'     If UserID = KAdiscountID And Val(Replace(txtTotal.Text, ",", "")) >= 1000 Then
      If Val(Replace(txtTotal.Text, ",", "")) >= 5000 Then
        txtCustName.Text = cboKA.List(cboKA.ListIndex)
        txtCustID.Text = cboKA.ItemData(cboKA.ListIndex)
        cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(txtCustID.Text) & ";")
        txtCusType.Text = "KEY ACCOUNTS"
        lblCustomer.Caption = "Key Accounts Cust."
        If OptCash.Value = True Then
            chkWholesaleCash.Value = 1
            txtPayScheme.Text = ""
            txtTerms.Text = ""
            txtNPays.Text = ""
            txtBillAmt.Text = ""
            txtPayStart.Text = ""
            lblTerms.Caption = ""
        Else
            chkWholesaleCash.Value = 0
            txtNPays.Text = 1
            txtPayStart.Text = txtDate.Text
        End If
    
        Ans = MsgBox("Do you want to apply Key Accounts Discount?", vbQuestion + vbYesNo, "Discount")
        If Ans = vbYes Then     'USE default Wholesale Price ver16.9.1
'            TtlOil = 0: TtlTire = 0: NonOilTire = 0
'            'PctDisc = IIf(Val(Replace(txtTotal.Text, ",", "")) >= 10000, 1.03, 1.05)
            With flxOrd
                For i = 1 To .Rows - 1
                    'If Wholesale Price Exist
                    If Val(.TextMatrix(i, 16)) > 0 Then
                        
'                         ApplyWholeSalePriceInRetail = True
                         ApplyKeyAcctsPriceInRetail = True
                         Qry = "SELECT SPAREPARTS.KAMAXDISCOUNT From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID =" & .TextMatrix(i, 10) & ""
                         myDscnt = Val(GetFieldValue(Qry)) / 100
                         Qry = "SELECT SPAREPARTS.KEYACCTS From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID =" & .TextMatrix(i, 10) & ""
                         myDscntLess = Val(GetFieldValue(Qry)) * myDscnt
                         
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 16)) - myDscntLess, "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 16), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
'                         WDisc = Int(((Val(Replace(.TextMatrix(i, 13), ",", "")) - Val(Replace(.TextMatrix(i, 16), ",", ""))) / Val(Replace(.TextMatrix(i, 13), ",", ""))) * 100)
                         
                         Qry = "Select KAMAXDISCOUNT From spareparts Where SPAREPARTID = " & .TextMatrix(i, 10) & ""
                         .TextMatrix(i, 6) = GetFieldValue(Qry) & "%"   'KADisc & "%"
                        If .TextMatrix(i, 15) = "OIL" Then
                            TtlOil = TtlOil + 1
                        ElseIf InStr(.TextMatrix(i, 15), "TIRE") > 0 Then
                            TtlTire = TtlTire + 1
                        Else
                            NonOilTire = NonOilTire + 1
                        End If
                    End If
                Next i
'                If TtlOil >= 1 And NonOilTire >= 1 Then 'to avail discount for oil, Purchase of 5k exclusive Oil only
'                    For i = 1 To .Rows - 1
'                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
'                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
'                         .TextMatrix(i, 6) = ""
'                         .TextMatrix(i, 6) = GetFieldValue(Qry) & "%"  'WDisc & "%"
'                    Next i
'                    MsgBox "Please use single invoice worth 5,000+ exclusive OIL Product to avail discount", vbInformation + vbOKOnly, "OIL Discount"
'                ElseIf TtlTire >= 1 And NonOilTire >= 1 Then 'to avail discount for Tires, Purchase of 5k exclusive Tires&Tube only
'                    For i = 1 To .Rows - 1
'                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
'                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
'                         .TextMatrix(i, 6) = ""
'                         .TextMatrix(i, 6) = GetFieldValue(Qry) & "%"  'WDisc & "%"
'                    Next i
'                    MsgBox "Please use single invoice worth 5,000+ exclusive TIRES & TUBES Product to avail discount", vbInformation + vbOKOnly, "TIRES & TUBE Discount"
'                End If
                        
                PrevTtl = Val(Replace(txtTotal.Text, ",", ""))
                Call ComputeTotal
            End With
            If PrevTtl > Val(Replace(txtTotal.Text, ",", "")) Then
                lblRef.Caption = "NOTE:"
                txtRefNo.Text = "Waiting for Terms ..."
                txtPctDisc.Text = Format(PrevTtl - Val(Replace(txtTotal.Text, ",", "")), "#,##0.00")

            End If
        End If
    ElseIf UserID <> KAdiscountID Then
        MsgBox "You are not Allowed to Key Accounts Discount!", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    ElseIf Val(Replace(txtTotal.Text, ",", "")) < 5000 Then
        MsgBox "Not Allowed to Key Accounts Discount. The Total Amount is Less than 5000!. ", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
    'mySQL
    cmdRemoveWS.Visible = True
    Call cmdKACancel_Click
    If OptOnAcct.Value = True And txtPayScheme.Text = "" Then
        lblTerms.Caption = "No Terms Applied"
    ElseIf OptOnAcct.Value = True And txtPayScheme.Text <> "" Then
        Ans = MsgBox("Do you want to use existing " & lblTerms.Caption & "?", vbQuestion + vbYesNo, "Terms")
        If Ans = vbNo Then
            txtPayScheme.Text = ""
            txtTerms.Text = ""
            txtNPays.Text = ""
            txtBillAmt.Text = ""
            txtPayStart.Text = ""
            lblTerms.Caption = "No Terms Applied"
        End If
    End If
    InsertPO.isKeyAcc = True
 InsertPO.Show vbModal
 
End Sub

Private Sub cmdPCCancel_Click()
    flxOrd.Enabled = True
    Frame2.Enabled = True
    tabPromoCard.Visible = False
    tabPromoCard.Top = 13000
    cmdTrans(3).Enabled = True
    cmdTrans(4).Enabled = True
End Sub

Private Sub cmdPCOK_Click()
Dim myValidation As String, myWholesale As String

If Val(txtGrandTotal.Text) > 0 Then
'    myValidation = ""
'    If Trim(txtPCCustomer.Text) <> "" Then
'        If cboPCPointsType.Text = "CUSTOMER POINTS" Then
'            myValidation = GetFieldValue("Select CUSTOMERNO From custpromocardinfo Where CUSTOMERNO = '" & txtPCCNo.Text & "' And INSTR(CARDTYPE,'MEMBERSHIP CARD') > 0 ")
'        Else
'            myValidation = GetFieldValue("Select CUSTOMERNO From custpromocardinfo Where CUSTOMERNO = '" & txtPCCNo.Text & "' And INSTR(CARDTYPE,'CLUB CARD') > 0 ")
'        End If
'        If myValidation <> "" Then
'            txtCustName.Text = txtPCCustomer.Text
'            txtCustID.Text = myCardCustomerID
'            cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(txtCustID.Text) & ";")
'            myCardPointsType = Trim(cboPCPointsType.Text) & "/EARN/CASH"
'            flxOrd.Enabled = True
'            Frame2.Enabled = True
'            tabPromoCard.Visible = False
'            tabPromoCard.Top = 13000
'            cmdTrans(3).Enabled = True
'            cmdTrans(4).Enabled = True
'            lblRef.Caption = "CARD Number"
'            txtRefNo.Text = myCardNo
'            cmdRemoveWS.Visible = True
'        Else
'            MsgBox "Invalid Points Type!.", vbInformation + vbOKOnly, "Validation"
'            Exit Sub
'        End If
        txtCustName.Text = txtPCCustomer.Text
        txtCustID.Text = myCardCustomerID
        cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(txtCustID.Text) & ";")
        myCardPointsType = Trim(cboPCPointsType.Text) & "/EARN/CASH"
        
        If Replace(txtTotal.Text, ",", "") >= 5000 Then
            myWholesale = GetFieldValue("Select CTYPE From CUSTOMER Where CUSTOMERNAME = '" & txtPCCustomer.Text & "' and CUSTOMERID = " & Val(txtCustID.Text) & " ")
            
            If myWholesale = "WHOLESALE" Then
                OptCash.Value = True
                Call ComputeWholesaleDiscount
                myCardPointsType = Trim(cboPCPointsType.Text) & "/EARN/CASH/WHOLESALE"
            Else
                MsgBox "This Customer is not Allowed to Wholesale!", vbInformation + vbOKOnly, "VALIDATION"
            End If
        End If
        flxOrd.Enabled = True
        Frame2.Enabled = True
        tabPromoCard.Visible = False
        tabPromoCard.Top = 13000
        cmdTrans(3).Enabled = True
        cmdTrans(4).Enabled = True
        lblRef.Caption = "CARD Number"
        txtRefNo.Text = myCardNo
        cmdRemoveWS.Visible = True
'    End If
Else
    MsgBox "Your Grand Total is " & Val(txtGrandTotal.Text) & ", Please Select a Partnumber first!", vbInformation + vbOKOnly, "Validation"
    Exit Sub
End If
End Sub

Private Sub cmdPO_Click()
    If rsParts.State = adStateClosed Then Timer1.Enabled = True
    InsertPO.Show vbModal
    
End Sub

Private Sub cmdRemoveCC_Click()
Ans = MsgBox("Do You Want to Remove Bank Charge?", vbQuestion + vbYesNo, "Validation")
If Ans = vbNo Then Exit Sub

For i = 1 To flxOrd.Rows - 1
    flxOrd.TextMatrix(i, 7) = Format(flxOrd.TextMatrix(i, 13), "#,##0.00")
Next i
Call ComputeTotal
cmdCreditCard.Visible = True
cmdRemoveCC.Visible = False
End Sub

Private Sub cmdRemoveWS_Click()
    Dim Ans As Integer, i As Integer
    
    If flxOrd.Rows = 2 And flxOrd.TextMatrix(1, 10) = "" Then
        'No Amount to Change Remove without prompting user
        txtCustName.Text = ""
        txtCustID.Text = ""
        txtCusType.Text = ""
        cboAddress.Text = ""
        If lblRef.Caption = "CARD Number" Then
            myCardNo = ""
            myCardCustomerID = 0
            myCardPointsType = ""
        End If
        cmdRemoveWS.Visible = False
        Exit Sub
    Else
        Ans = MsgBox("Do you want to remove Customer " & txtCustName.Text & "?" & vbCrLf & _
            "The prices of all items ordered will change to original SRP!", vbQuestion + vbYesNo, "Remove Customer")
        If Ans = vbNo Then Exit Sub
    End If
    If txtRefNo.Text = "WHOLESALE DISCOUNT" Then
        lblRef.Caption = "Reference No."
    End If
    txtRefNo.Text = ""
    txtCustName.Text = ""
    'lblRemarks.Caption = ""
    cboAddress.Text = ""
    txtCustID.Text = ""
    txtCusType.Text = ""
    cmdRemoveWS.Visible = False
    txtPayScheme.Text = ""
    txtTerms.Text = ""
    txtNPays.Text = ""
    txtBillAmt.Text = ""
    txtPayStart.Text = ""
    If lblRef.Caption = "CARD Number" Then
        txtRefNo.Text = ""
        lblRef.Caption = "Reference No."
        myCardNo = ""
        myCardCustomerID = 0
        myCardPointsType = ""
    End If
    With flxOrd
        For i = 1 To .Rows - 1
            .TextMatrix(i, 7) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")), "#,##0.00")
            .TextMatrix(i, 6) = ""
        Next i
    End With
    Call ComputeTotal
End Sub

Private Sub cmdReq_Click()
    Dim rsReq As New ADODB.Recordset
    Dim rsReqDtls As New ADODB.Recordset
    Dim MyReq As String, Qry As String
    Dim Ans As Integer, ReqID As Long
    
    If flxOrd.Rows >= 2 And flxOrd.TextMatrix(flxOrd.Rows - 1, 10) <> "" Then
        Ans = MsgBox("All Spareparts  listed in the Order Slip will be removed!" & vbCrLf & "Do you want to Continue?", vbQuestion + vbYesNo)
        If Ans = vbNo Then Exit Sub
    End If
    MyReq = InputBox("Please enter the Requisition Number:", "Req. No.")
    If MyReq = "" Then Exit Sub
    'Qry = "SELECT REQUISITION.REQNO, REQUISITION.JOBNO, REQUISITION.CUSTOMER, REQUISITION.REQID, REQUISITION.JOBORDERID, REQUISITION.CUSTOMERID " & _
          "From Requisition WHERE (((REQUISITION.REQNO)='" & MyReq & "'));"
    'mySQL
    Qry = "SELECT REQUISITION.REQNO, REQUISITION.JOBNO, REQUISITION.CUSTOMER, REQUISITION.REQID, REQUISITION.JOBORDERID, REQUISITION.CUSTOMERID " & _
          "From Requisition WHERE (((REQUISITION.REQNO)='" & MyReq & "'));"
    
    With rsReq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ReqID = !ReqID
            lblRef.Caption = "Requisition No."
            txtRefNo.Text = !REQNO
            lblRemarks.Caption = "Job Order No. " & !JobNo
            txtCustName.Text = !CUSTOMER
            lblCustomer.Caption = "Service Customer"
            txtJOid.Text = !JOBORDERID
            txtCustID.Text = !CUSTOMERID
        Else
            ReqID = 0
        End If
        .Close
    End With
    If ReqID = 0 Then
        MsgBox "Requisition No. " & MyReq & " does not exist!", vbExclamation + vbOKOnly, "Requisition"
        Exit Sub
    End If
    If Val(txtCustID.Text) > 0 Then
        cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(txtCustID.Text) & ";")
    End If
    flxOrd.Rows = 1
    'Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST " & _
          "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.REQID)=" & ReqID & "));"
    'mySQL
    
    '**Update 05/11/2010
    Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.PRICE*SPSORDERDETAIL.QTY AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST " & _
          "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE SPSORDERDETAIL.REQID=" & ReqID & ";"
    
    With rsReqDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem vbTab & !QtyOnHand & vbTab & !Qty & vbTab & IIf(IsNull(!Unit), "", !Unit) & vbTab & !PartNumber & vbTab & !Description & vbTab & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparepartID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & !MaxDiscount & vbTab & !Price & vbTab & !Cost
                .MoveNext
            Loop
        End If
        .Close
    End With
    txtReqID.Text = ReqID
    Call ComputeTotal
    If Val(txtCustID.Text) > 0 Then
        Call GetCustomerDefaultTerm
        lblCustomer.Caption = "On-Acct. Customer"
    End If
    If rsParts.State = adStateClosed Then Timer1.Enabled = True
    Call LockData(False)
End Sub

Private Sub GetCustomerDefaultTerm()
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, Gtotal As Currency
    
    'Saving On Account Job Order to On Account Terms
    'Get Terms
    'Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
    Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
          "WHERE ONACCTERMS.STATUS='Default' AND ONACCTERMS.CUSTOMERID=" & txtCustID.Text & ";"
    
    With rsTerm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtTerms.Text = !Terms
            txtPayScheme.Text = !PayScheme
            txtNPays.Text = !Npays
        Else
            txtTerms.Text = 30
            txtPayScheme.Text = "MONTHLY"
            txtNPays.Text = 1
        End If
        .Close
    End With
    txtPayStart.Text = StartPay(txtPayScheme.Text)
    Gtotal = Val(Replace(txtGrandTotal.Text, ",", ""))
    If Val(txtNPays.Text) > 0 Then
        txtBillAmt.Text = Gtotal / Val(txtNPays.Text)
    Else
        txtBillAmt.Text = Gtotal
    End If
End Sub
Function StartPay(S As String) As Date
    Dim iPS As Integer, sPay As Date, MyDate As Date
    
    MyDate = txtDate.Text
    If S = "WEEKLY" Then iPS = 7
    If S = "BI-MONTHLY" Then iPS = 15
    If S = "MONTHLY" Then iPS = 30
    If iPS = 7 Then
        If Day(MyDate) < 5 Then
            sPay = DateSerial(Year(MyDate), Month(MyDate), 7)
        ElseIf Day(MyDate) < 12 Then
            sPay = DateSerial(Year(MyDate), Month(MyDate), 14)
        ElseIf Day(MyDate) < 19 Then
            sPay = DateSerial(Year(MyDate), Month(MyDate), 21)
        ElseIf Day(MyDate) < 28 Then
            sPay = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, MyDate)), Month(DateAdd("m", 1, MyDate)), 1))
        Else
            sPay = DateAdd("m", 1, DateSerial(Year(MyDate), Month(MyDate), 7))
        End If
        StartPay = sPay
    ElseIf iPS = 15 Then
        StartPay = DateAdd("d", iPS, MyDate)
    ElseIf iPS = 30 Then
        StartPay = DateAdd("m", 1, MyDate)
    End If
End Function

Private Sub cmdRPCancel_Click()
    flxOrd.Enabled = True
    Frame2.Enabled = True
    tabPromoCard.Visible = False
    tabPromoCard.Top = 13000
    cmdTrans(3).Enabled = True
    cmdTrans(4).Enabled = True
End Sub

Private Sub cmdRPOK_Click()
If Trim(txtRPCustomer.Text) <> "" Then
    If Val(txtGrandTotal.Text) > 0 Then
        If Val(txtRPCardPoints.Text) >= Val(txtGrandTotal.Text) Then
            txtCustName.Text = txtRPCustomer.Text
            txtCustID.Text = myCardCustomerID
            cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(txtCustID.Text) & ";")
            myCardPointsType = Trim(cboRPPointsType.Text) & "/REDEEM/CASH"
            flxOrd.Enabled = True
            Frame2.Enabled = True
            tabPromoCard.Visible = False
            tabPromoCard.Top = 13000
            cmdTrans(3).Enabled = True
            cmdTrans(4).Enabled = True
            lblRef.Caption = "CARD Number"
            txtRefNo.Text = myCardNo
            cmdRemoveWS.Visible = True
        Else
            MsgBox "Insufficient Card Points", vbInformation + vbOKOnly, "Validation"
            Exit Sub
        End If
    Else
        MsgBox "Your Grand Total is " & Val(txtGrandTotal.Text) & ", Please Select a Partnumber first!", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
End If
End Sub

Private Sub cmdSubCreditCard_Click()
flxOrd.Enabled = False
Frame2.Enabled = False
cmdTrans(3).Enabled = False
cmdTrans(4).Enabled = False
txtBankCharge.Text = "0"
tabCommandCC.Visible = False
With ssCreditTab
    .Top = (Me.Height - .Height) / 2
    .Left = (Me.Width - .Width) / 2
    .Visible = True
    .SetFocus
End With
'MsgBox "Not Yet Available", vbInformation + vbOKOnly, "Validation" cmdSubCreditCard
End Sub

Private Sub cmdSubPromoCard_Click()

flxOrd.Enabled = False
Frame2.Enabled = False
cmdTrans(3).Enabled = False
cmdTrans(4).Enabled = False

txtPCCustomer.Text = ""
txtPCCNo.Text = ""
txtPCCPoints.Text = "0"
cboPCPointsType.Text = ""
cboPCPointsType.Clear
cboPCPointsType.AddItem "CUSTOMER POINTS"
cboPCPointsType.AddItem "CLUB POINTS"

txtRPCustomer.Text = ""
txtRPCardNo.Text = ""
txtRPCardPoints.Text = "0"
cboRPPointsType.Text = ""
cboRPPointsType.Clear
cboRPPointsType.AddItem "CUSTOMER POINTS"
cboRPPointsType.AddItem "CLUB POINTS"

tabCommandCC.Visible = False

With tabPromoCard
    .Top = (Me.Height - .Height) / 2
    .Left = (Me.Width - .Width) / 2
    .Visible = True
    .SetFocus
End With
'MsgBox "Not Yet Available", vbInformation + vbOKOnly, "Validation"
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 12) As String, Qry As String
    
    On Error GoTo LocalErr
    Select Case Index
        Case 0  'New
            Call ClearEntries
            Call LockData(False)
            'txtDate.Text = Date
            txtOrderno.Text = "<Auto>"
            AddNew = True
            Call DisplayMainButton(1)
            cmdCancel.Visible = False
            'cmdCashier.Visible = False
            cmdWareHouse.Visible = False
            If rsParts.State = adStateOpen Then rsParts.Close
            Timer1.Enabled = True
            flxPND.Cols = 0
            flxOrd.SetFocus
            flxOrd.Col = 4
        Case 1  'Edit
            If txtSPSorderID.Text = "" Then
                MsgBox "No available Sales Order to edit", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If txtInvNo.Text <> "" Then
                MsgBox "Unable to edit Order# " & txtOrderno.Text & " because an Invoice was already issued.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If InStr(lblRemarks.Caption, "CANCELLED") > 0 Then
                MsgBox "Unable to edit Sales Order " & txtOrderno.Text & " because it has already been Cancelled.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If txtRefNo.Text = "WHOLESALE DISCOUNT" Then
                 Qry = "SELECT SETTINGS.REMARKS From SETTINGS " & _
                      "WHERE SETTINGS.DESCRIPTION = 'BM DISCOUNT' AND SETTINGS.VALUE = '8/01/2016';"
                WSdiscountID = Val(GetFieldValue(Qry))
                If UserID <> WSdiscountID Then
                    MsgBox "Sorry, You are not allowed to EDIT this sales order" & vbCrLf & "Please contact MIS Helpdesk for more information.", vbOKOnly + vbExclamation, "Discounted Sales Order"
                    Exit Sub
                End If
            End If
            If rsParts.State = adStateOpen Then rsParts.Close
            Timer1.Enabled = True
            flxPND.Cols = 0
            Call DisplayMainButton(1)
            AddNew = False
        Case 2  'Exit
            'MotoRep.tbMain.Visible = True
            'ProperExit = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
            
        Case 3  'Save
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            myMsg = ValidateSalesOrder
            If InsertPO.isKeyAcc Then
                If txtRefNo = "" Then
                 MsgBox "Please Insert Terms", vbCritical + vbOKOnly, "Validation Error"
                 Exit Sub
                End If
            End If
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save order?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Save Spare Parts Invoice
            Call SaveSalesOrder
            Ans = MsgBox("Sales Order Successfully Saved!" & vbCrLf & vbCrLf & "Process Another Order?", vbQuestion + vbYesNo, "Order")
            If Ans = vbNo Then
                Call DisplayMainButton(2)
                cmdCancel.Visible = True
                'cmdCashier.Visible = True
                cmdWareHouse.Visible = True
                cmdRemoveCC.Visible = False
                If UserType <> "Administrator" Then
                    Call DisplayPendingOrders
                End If
            Else
                Call cmdTrans_Click(0)
            End If
        Case 4  'Cancel
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            Ans = MsgBox("Do you want to cancel Order?", vbQuestion + vbYesNo)
            If Ans = vbNo Then Exit Sub
            Call DisplayMainButton(2)
            Call ClearEntries
            Ans = MsgBox("Process Another Order?", vbQuestion + vbYesNo, "Order")
            If Ans = vbNo Then
                'MotoRep.tbMain.Visible = True
                'MotoRep.imgBG.Visible = True
                MotoRep.toolBarButton.Visible = True
                Unload Me
                'LogIn.Show
            Else
                cmdCancel.Visible = True
                'cmdCashier.Visible = True
                cmdWareHouse.Visible = True
                If UserType <> "Administrator" Then
                    Call DisplayPendingOrders
                End If
                Call LockData(True)
                'Call cmdTrans_Click(0)
            End If
    End Select
    Exit Sub

LocalErr:
    MsgBox "Error Code: " & CStr(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub

Private Sub cmdTrans_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
    If KeyCode = vbKeyF9 And Not JustScroll Then Call cmdCashier_Click 'Key Accounts
End Sub
'-------------Additional Verification for Adding Discount for Wholesale and Key Accounts 04/20/2024------
Private Sub cmdVerify_Click()
    Dim rsFinalRec As New ADODB.Recordset
    Dim rsKA As New ADODB.Recordset
    
'    Qry = "SELECT pcuser.USERNAME,pcuser.Password From pcuser Where pcuser.PASSWORD = '" & txtCode.Text & "';"
    Qry = "SELECT pcuser.USERNAME,pcuser.PASSWORD FROM pcuser INNER JOIN employees ON (pcuser.EMPID = employees.EMPID) WHERE employees.PCUSER = 'MANAGER' AND pcuser.PASSWORD = '" & txtCode.Text & "';"
        With rsFinalRec
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount <= 0 Then
                MsgBox "Invalid Code!", vbExclamation + vbOKOnly, "Access Denied!"
                tabKA.Visible = False
                Exit Sub
            Else
            tabFinal.Visible = False
            End If
            .Close
        End With
        
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER WHERE CUSTOMER.CTYPE='KEY ACCOUNTS' ORDER BY CUSTOMER.CUSTOMERNAME;"
    cboKA.Clear
    With rsKA
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboKA.AddItem !CustomerName
                cboKA.ItemData(cboKA.ListCount - 1) = !CUSTOMERID
                .MoveNext
            Loop
        End If
        .Close
    End With

    flxOrd.Enabled = False
    Frame2.Enabled = False
    cmdTrans(3).Enabled = False
    cmdTrans(4).Enabled = False
    txtRefNo.Locked = True

    With tabKA
        .Caption = ""
        .Top = (Me.Height - .Height) / 2
        .Left = (Me.Width - .Width) / 2
        .Visible = True
        .SetFocus
    End With

End Sub

Private Sub cmdWareHouse_Click()
    Dim i As Integer, vSkip As Boolean, eMsg As String
    
    'Validate items if all are found
    On Error GoTo LocalErr
    
    With flxOrd
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "" Then
                vSkip = True
                Exit For
            Else
                vSkip = False
            End If
        Next i
    End With
    If Not vSkip Then
        MsgBox "No available items for Picking!", vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    eMsg = "Updating Destination of Sales Order"
    'LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.DESTINATION = 'WAREHOUSE' " & _
                   "WHERE (((SPSALESORDER.SPSORDERID)=" & txtSPSorderID.Text & "));"
    LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.DESTINATION = 'WAREHOUSE' " & _
                   "WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    MsgBox "Sales Order can now be viewed in Warehouse!", vbInformation + vbOKOnly, "Updated in Warehouse"
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub

Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xRow As Integer, xBal As Integer, i As Integer
    
    'Clear FlxStock
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyID & ")) ORDER BY SPITR.TRANSDATE;"
        
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
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
                
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

Function GetSalesQty(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String, StartDate As Date, EndDate As Date
    Dim MySales As Long, MyRet As Long
    
    StartDate = DateSerial(Year(DateAdd("m", -3, Date)), Month(DateAdd("m", -3, Date)), 1)
    EndDate = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), Day(DateAdd("d", -1, DateSerial(Year(Date), Month(Date), 1))))
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
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
    'Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>='" & Format(StartDate, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(EndDate, "yyyy-mm-dd") & "'));"
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

Private Sub cmdWholesale_Click()
    Dim rsWS As New ADODB.Recordset
    
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER WHERE (((CUSTOMER.CTYPE)='WHOLESALE')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER WHERE CUSTOMER.CTYPE='WHOLESALE' ORDER BY CUSTOMER.CUSTOMERNAME;"
    
    cboWS.Clear
    With rsWS
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboWS.AddItem !CustomerName
                cboWS.ItemData(cboWS.ListCount - 1) = !CUSTOMERID
                .MoveNext
            Loop
        End If
        .Close
    End With
    flxOrd.Enabled = False
    Frame2.Enabled = False
    cmdTrans(3).Enabled = False
    cmdTrans(4).Enabled = False
    txtRefNo.Locked = True
    With tabWS
        .Caption = ""
        .Top = (Me.Height - .Height) / 2
        .Left = (Me.Width - .Width) / 2
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub cmdWSCancel_Click()
    flxOrd.Enabled = True
    Frame2.Enabled = True
    tabWS.Visible = False
    tabWS.Top = 13000
    cmdTrans(3).Enabled = True
    cmdTrans(4).Enabled = True
End Sub

Private Sub cmdWSOK_Click()
    Dim rsDisc As New ADODB.Recordset
    Dim Qry As String, Ans As Integer, i As Integer, PctDisc As Currency
    Dim NewPrc As Currency, WDisc As Currency, PrevTtl As Currency
    Dim myDscnt As Currency, myDscntLess As Currency
    Dim TtlOil As Integer, TtlTire As Integer, NonOilTire As Integer
    
    If cboWS.ListIndex = -1 Then
        MsgBox "Please select a customer from the list!", vbExclamation + vbOKOnly, "Wholesale"
        Exit Sub
    End If
    
    'Validate user if allowed to Process
     Qry = "SELECT SETTINGS.REMARKS From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION = 'BM DISCOUNT' AND SETTINGS.VALUE = '8/01/2016';"
     WSdiscountID = Val(GetFieldValue(Qry))
     
     If UserID = WSdiscountID And Val(Replace(txtTotal.Text, ",", "")) >= 5000 Then
    
'     If UserID = WSdiscountID And Val(Replace(txtTotal.Text, ",", "")) >= 5000 And (OptCash.Value = True Or (OptOnAcct.Value = True And InStr(txtCustName.Text, "GAS ISLAND") > 0)) Then
'
        txtCustName.Text = cboWS.List(cboWS.ListIndex)
        txtCustID.Text = cboWS.ItemData(cboWS.ListIndex)
        cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(txtCustID.Text) & ";")
        txtCusType.Text = "WHOLESALE"
        lblCustomer.Caption = "Wholesale Cust."
        If OptCash.Value = True Then
            chkWholesaleCash.Value = 1
            txtPayScheme.Text = ""
            txtTerms.Text = ""
            txtNPays.Text = ""
            txtBillAmt.Text = ""
            txtPayStart.Text = ""
            lblTerms.Caption = ""
        Else
            chkWholesaleCash.Value = 0
        End If
    
        Ans = MsgBox("Do you want to apply Wholesale Discount?", vbQuestion + vbYesNo, "Discount")
        If Ans = vbYes Then     'USE default Wholesale Price ver16.9.1
            TtlOil = 0: TtlTire = 0: NonOilTire = 0
            'PctDisc = IIf(Val(Replace(txtTotal.Text, ",", "")) >= 10000, 1.03, 1.05)
            With flxOrd
                For i = 1 To .Rows - 1
                    'If Wholesale Price Exist
                    If Val(.TextMatrix(i, 16)) > 0 Then
'                        'Display Discounted Price Before ver16.9.1
'                        NewPrc = Int(Val(.TextMatrix(i, 16)) * PctDisc)
'                        If NewPrc < Val(.TextMatrix(i, 13)) Then
'                            .TextMatrix(i, 7) = Format(NewPrc, "#,##0.00")
'                            .TextMatrix(i, 8) = Format(NewPrc * Val(.TextMatrix(i, 2)), "#,##0.00")
'                            WDisc = Int(((Val(.TextMatrix(i, 13)) - NewPrc) / Val(.TextMatrix(i, 13))) * 100)
'                            .TextMatrix(i, 6) = WDisc & "%"
'                        End If
                        
                         ApplyWholeSalePriceInRetail = True
                         myDscnt = Val(.TextMatrix(i, 17)) / 100
                         myDscntLess = Val(.TextMatrix(i, 16)) * myDscnt
                         
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 16)) - myDscntLess, "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 16), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
'                         WDisc = Int(((Val(Replace(.TextMatrix(i, 13), ",", "")) - Val(Replace(.TextMatrix(i, 16), ",", ""))) / Val(Replace(.TextMatrix(i, 13), ",", ""))) * 100)
                         
                         Qry = "Select WSMAXDISCOUNT From spareparts Where SPAREPARTID = " & .TextMatrix(i, 10) & ""
                         .TextMatrix(i, 6) = GetFieldValue(Qry) & "%"   'WDisc & "%"
                        If .TextMatrix(i, 15) = "OIL" Then
                            TtlOil = TtlOil + 1
                        ElseIf InStr(.TextMatrix(i, 15), "TIRE") > 0 Then
                            TtlTire = TtlTire + 1
                        Else
                            NonOilTire = NonOilTire + 1
                        End If
                    End If
                Next i
                If TtlOil >= 1 And NonOilTire >= 1 Then 'to avail discount for oil, Purchase of 5k exclusive Oil only
                    For i = 1 To .Rows - 1
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                         .TextMatrix(i, 6) = ""
'                         Qry = "Select WSMAXDISCOUNT From spareparts Where SPAREPARTID = " & .TextMatrix(i, 10) & ""
'                         .TextMatrix(i, 6) = GetFieldValue(Qry) & "%"  'WDisc & "%"
                    Next i
                    MsgBox "Please use single invoice worth 5,000+ exclusive OIL Product to avail discount", vbInformation + vbOKOnly, "OIL Discount"
                ElseIf TtlTire >= 1 And NonOilTire >= 1 Then 'to avail discount for Tires, Purchase of 5k exclusive Tires&Tube only
                    For i = 1 To .Rows - 1
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                         .TextMatrix(i, 6) = ""
'                         Qry = "Select WSMAXDISCOUNT From spareparts Where SPAREPARTID = " & .TextMatrix(i, 10) & ""
'                         .TextMatrix(i, 6) = GetFieldValue(Qry) & "%"  'WDisc & "%"
                    Next i
                    MsgBox "Please use single invoice worth 5,000+ exclusive TIRES & TUBES Product to avail discount", vbInformation + vbOKOnly, "TIRES & TUBE Discount"
                End If
                        
                PrevTtl = Val(Replace(txtTotal.Text, ",", ""))
                Call ComputeTotal
            End With
            If PrevTtl > Val(Replace(txtTotal.Text, ",", "")) Then
                lblRef.Caption = "NOTE:"
                txtRefNo.Text = "WHOLESALE DISCOUNT"
                txtPctDisc.Text = Format(PrevTtl - Val(Replace(txtTotal.Text, ",", "")), "#,##0.00")
'            Else
'                lblRef.Caption = "NOTE:"
'                txtRefNo.Text = "WHOLESALE DISCOUNT"
'                txtPctDisc.Text = Format(PrevTtl - Val(Replace(txtTotal.Text, ",", "")), "#,##0.00")
            End If
        End If
    ElseIf UserID <> WSdiscountID Then
        MsgBox "You are not Allowed to Wholesale Discount!", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    ElseIf Val(Replace(txtTotal.Text, ",", "")) < 5000 Then
        MsgBox "Not Allowed to Wholesale Discount. The Total Amount is Less than 5000!. ", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
    'mySQL
    cmdRemoveWS.Visible = True
    Call cmdWSCancel_Click
    If OptOnAcct.Value = True And txtPayScheme.Text = "" Then
        lblTerms.Caption = "No Terms Applied"
    ElseIf OptOnAcct.Value = True And txtPayScheme.Text <> "" Then
        Ans = MsgBox("Do you want to use existing " & lblTerms.Caption & "?", vbQuestion + vbYesNo, "Terms")
        If Ans = vbNo Then
            txtPayScheme.Text = ""
            txtTerms.Text = ""
            txtNPays.Text = ""
            txtBillAmt.Text = ""
            txtPayStart.Text = ""
            lblTerms.Caption = "No Terms Applied"
        End If
    End If
'    If flxOrd.Rows = 2 And flxOrd.TextMatrix(1, 10) = "" Then
'        Exit Sub
'    Else
'        Ans = MsgBox("Do you want the existing orders to avail the wholesale discount?", vbQuestion + vbYesNo, "Wholesale")
'        If Ans = vbNo Then Exit Sub
'        Call ComputeWholesaleDiscount
'    End If
End Sub

Public Sub ComputeWholesaleDiscount()
    Dim PctDisc As Currency, DiscPrice As Currency, i As Integer, j As Integer, myVat As Single
    Dim myRem As String, Qry As String, TtlAmnt As Currency, TtlOrgPrc As Currency
    
'    With flxOrd
'        For j = 1 To .Rows - 1
'            If .TextMatrix(j, 10) <> "" Then
'                PctDisc = 0
'
'                Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(j, 10)) & "));"
'                PctDisc = Val(GetFieldValue(Qry))
'
'                'old code
'                'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
'                '    If Val(.TextMatrix(j, 10)) = Val(flxSPec.TextMatrix(i, 0)) Then
'                '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
'                '        Exit For
'                '    End If
'                'Next i
'
'                If PctDisc = 0 Then
'
'                    rsParts.MoveFirst
'                    rsParts.Find "[SPAREPARTID] = " + .TextMatrix(j, 10)
'                    If Not rsParts.EOF Then
'                        myRem = rsParts!Remarks
'                    Else
'                        myRem = ""
'                    End If
'                    If myRem <> "" Then
'                        For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
'                            If myRem = flxWDisc.TextMatrix(i, 0) Then
'                                PctDisc = Val(flxWDisc.TextMatrix(i, 1))
'                                Exit For
'                            End If
'                        Next i
'                    End If
'                    If PctDisc > 0 Then
'                        myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
'                        DiscPrice = (Val(.TextMatrix(j, 14)) * myVat) * PctDisc
'                         WDisc = Int(((Val(.TextMatrix(i, 13)) - Val(.TextMatrix(i, 16))) / Val(.TextMatrix(i, 13))) * 100)
'                         .TextMatrix(i, 6) = WDisc & "%"
'                        If InStr(myRem, "GENUINE") > 0 Then
'                            If DiscPrice < 10 Then
'                                DiscPrice = DiscPrice + 3
'                            ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
'                                DiscPrice = DiscPrice + 2
'                            Else
'                                If InStr(str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
'                            End If
'                            DiscPrice = Int(DiscPrice)
'                        Else
'                            If InStr(str(DiscPrice), ".") > 0 Then
'                                DiscPrice = DiscPrice + 1
'                                DiscPrice = Int(DiscPrice)
'                            End If
'                        End If
'                        flxOrd.TextMatrix(j, 7) = Format$(DiscPrice, "###,##0.00") 'Price
'                    Else
'                        flxOrd.TextMatrix(j, 7) = Format$(Val(.TextMatrix(j, 13)), "###,##0.00") 'Orig SRP
'                    End If
'                Else    'Display Amount from Exclusion List
'                    flxOrd.TextMatrix(j, 7) = Format$(PctDisc, "###,##0.00") 'Price
'                End If
'            End If
'        Next j
'    End With
        'Validate user if allowed to Process
     Qry = "SELECT SETTINGS.REMARKS From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION = 'BM DISCOUNT' AND SETTINGS.VALUE = '8/01/2016';"
     WSdiscountID = Val(GetFieldValue(Qry))
     If UserID = WSdiscountID And Val(Replace(txtTotal.Text, ",", "")) >= 5000 And (OptCash.Value = True Or (OptOnAcct.Value = True And InStr(txtCustName.Text, "GAS ISLAND") > 0)) Then
    
        Ans = MsgBox("Do you want to apply Wholesale Discount?", vbQuestion + vbYesNo, "Discount")
        If Ans = vbYes Then     'USE default Wholesale Price ver16.9.1
            TtlOil = 0: TtlTire = 0: NonOilTire = 0
            'PctDisc = IIf(Val(Replace(txtTotal.Text, ",", "")) >= 10000, 1.03, 1.05)
            With flxOrd
                For i = 1 To .Rows - 1
                    'If Wholesale Price Exist
                    If Val(.TextMatrix(i, 16)) > 0 Then
'                        'Display Discounted Price Before ver16.9.1
'                        NewPrc = Int(Val(.TextMatrix(i, 16)) * PctDisc)
'                        If NewPrc < Val(.TextMatrix(i, 13)) Then
'                            .TextMatrix(i, 7) = Format(NewPrc, "#,##0.00")
'                            .TextMatrix(i, 8) = Format(NewPrc * Val(.TextMatrix(i, 2)), "#,##0.00")
'                            WDisc = Int(((Val(.TextMatrix(i, 13)) - NewPrc) / Val(.TextMatrix(i, 13))) * 100)
'                            .TextMatrix(i, 6) = WDisc & "%"
'                        End If
                         ApplyWholeSalePriceInRetail = True
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 16)), "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 16), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                         WDisc = Int(((Val(Replace(.TextMatrix(i, 13), ",", "")) - Val(Replace(.TextMatrix(i, 16), ",", ""))) / Val(Replace(.TextMatrix(i, 13), ",", ""))) * 100)
                         .TextMatrix(i, 6) = WDisc & "%"
                        If .TextMatrix(i, 15) = "OIL" Then
                            TtlOil = TtlOil + 1
                        ElseIf InStr(.TextMatrix(i, 15), "TIRE") > 0 Then
                            TtlTire = TtlTire + 1
                        Else
                            NonOilTire = NonOilTire + 1
                        End If
                    End If
                Next i
                If TtlOil >= 1 And NonOilTire >= 1 Then 'to avail discount for oil, Purchase of 5k exclusive Oil only
                    For i = 1 To .Rows - 1
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                         .TextMatrix(i, 6) = ""
                    Next i
                    MsgBox "Please use single invoice worth 5,000+ exclusive OIL Product to avail discount", vbInformation + vbOKOnly, "OIL Discount"
                ElseIf TtlTire >= 1 And NonOilTire >= 1 Then 'to avail discount for Tires, Purchase of 5k exclusive Tires&Tube only
                    For i = 1 To .Rows - 1
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                         .TextMatrix(i, 6) = ""
                    Next i
                    MsgBox "Please use single invoice worth 5,000+ exclusive TIRES & TUBES Product to avail discount", vbInformation + vbOKOnly, "TIRES & TUBE Discount"
                End If
                        
                PrevTtl = Val(Replace(txtTotal.Text, ",", ""))
                Call ComputeTotal
            End With
            If PrevTtl > Val(Replace(txtTotal.Text, ",", "")) Then
                lblRef.Caption = "NOTE:"
                txtRefNo.Text = "WHOLESALE DISCOUNT"
                txtPctDisc.Text = Format(PrevTtl - Val(Replace(txtTotal.Text, ",", "")), "#,##0.00")
            End If
        End If
    Else
        TtlOrgPrc = 0
        TtlAmnt = 0
        With flxOrd
            For i = 1 To .Rows - 1
                TtlOrgPrc = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                TtlAmnt = Format(TtlAmnt + TtlOrgPrc, "#,##0.00")
            Next i
            
            If UserID <> WSdiscountID Then
                MsgBox "This User is not allowed to process Wholesale!", vbInformation + vbOKOnly, "Validation"
            ElseIf Val(Replace(TtlAmnt, ",", "")) < 5000 Then
                MsgBox "WHOLESALE IS NOT AVAILABLE! THE PRICES WILL BE BACK TO ORIGINAL PRICE", vbInformation + vbOKOnly, "Validation"
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 7) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")), "#,##0.00")
                    .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                    .TextMatrix(i, 6) = ""
                Next i
                txtRefNo.Text = ""
                lblRef.Caption = "Reference No."
            End If
        End With
    End If
    Call ComputeTotal
End Sub
Public Sub ComputeKeyAcctsDiscount()
    Dim PctDisc As Currency, DiscPrice As Currency, i As Integer, j As Integer, myVat As Single
    Dim myRem As String, Qry As String, TtlAmnt As Currency, TtlOrgPrc As Currency
    
    Qry = "SELECT SETTINGS.REMARKS From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION = 'BM DISCOUNT' AND SETTINGS.VALUE = '8/01/2016';"
     WSdiscountID = Val(GetFieldValue(Qry))
     If UserID = WSdiscountID And Val(Replace(txtTotal.Text, ",", "")) >= 5000 And (OptCash.Value = True Or (OptOnAcct.Value = True And InStr(txtCustName.Text, "GAS ISLAND") > 0)) Then
    
        Ans = MsgBox("Do you want to apply Wholesale Discount?", vbQuestion + vbYesNo, "Discount")
        If Ans = vbYes Then     'USE default Wholesale Price ver16.9.1
            TtlOil = 0: TtlTire = 0: NonOilTire = 0
            'PctDisc = IIf(Val(Replace(txtTotal.Text, ",", "")) >= 10000, 1.03, 1.05)
            With flxOrd
                For i = 1 To .Rows - 1
                    'If Wholesale Price Exist
                    If Val(.TextMatrix(i, 16)) > 0 Then
'                        'Display Discounted Price Before ver16.9.1
'                        NewPrc = Int(Val(.TextMatrix(i, 16)) * PctDisc)
'                        If NewPrc < Val(.TextMatrix(i, 13)) Then
'                            .TextMatrix(i, 7) = Format(NewPrc, "#,##0.00")
'                            .TextMatrix(i, 8) = Format(NewPrc * Val(.TextMatrix(i, 2)), "#,##0.00")
'                            WDisc = Int(((Val(.TextMatrix(i, 13)) - NewPrc) / Val(.TextMatrix(i, 13))) * 100)
'                            .TextMatrix(i, 6) = WDisc & "%"
'                        End If
                         ApplyKeyAcctsPriceInRetail = True
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 16)), "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 16), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                         WDisc = Int(((Val(Replace(.TextMatrix(i, 13), ",", "")) - Val(Replace(.TextMatrix(i, 16), ",", ""))) / Val(Replace(.TextMatrix(i, 13), ",", ""))) * 100)
                         .TextMatrix(i, 6) = WDisc & "%"
                        If .TextMatrix(i, 15) = "OIL" Then
                            TtlOil = TtlOil + 1
                        ElseIf InStr(.TextMatrix(i, 15), "TIRE") > 0 Then
                            TtlTire = TtlTire + 1
                        Else
                            NonOilTire = NonOilTire + 1
                        End If
                    End If
                Next i
                If TtlOil >= 1 And NonOilTire >= 1 Then 'to avail discount for oil, Purchase of 5k exclusive Oil only
                    For i = 1 To .Rows - 1
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                         .TextMatrix(i, 6) = ""
                    Next i
                    MsgBox "Please use single invoice worth 5,000+ exclusive OIL Product to avail discount", vbInformation + vbOKOnly, "OIL Discount"
                ElseIf TtlTire >= 1 And NonOilTire >= 1 Then 'to avail discount for Tires, Purchase of 5k exclusive Tires&Tube only
                    For i = 1 To .Rows - 1
                         .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 13)), "#,##0.00")
                         .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                         .TextMatrix(i, 6) = ""
                    Next i
                    MsgBox "Please use single invoice worth 5,000+ exclusive TIRES & TUBES Product to avail discount", vbInformation + vbOKOnly, "TIRES & TUBE Discount"
                End If
                        
                PrevTtl = Val(Replace(txtTotal.Text, ",", ""))
                Call ComputeTotal
            End With
            If PrevTtl > Val(Replace(txtTotal.Text, ",", "")) Then
                lblRef.Caption = "NOTE:"
                txtRefNo.Text = "KEY ACCOUNTS DISCOUNT"
                txtPctDisc.Text = Format(PrevTtl - Val(Replace(txtTotal.Text, ",", "")), "#,##0.00")
            End If
        End If
    Else
        TtlOrgPrc = 0
        TtlAmnt = 0
        With flxOrd
            For i = 1 To .Rows - 1
                TtlOrgPrc = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                TtlAmnt = Format(TtlAmnt + TtlOrgPrc, "#,##0.00")
            Next i
            
            If UserID <> WSdiscountID Then
                MsgBox "This User is not allowed to process Kay Accounts!", vbInformation + vbOKOnly, "Validation"
            ElseIf Val(Replace(TtlAmnt, ",", "")) < 5000 Then
                MsgBox "KEY ACCOUNTS IS NOT AVAILABLE! THE PRICES WILL BE BACK TO ORIGINAL PRICE", vbInformation + vbOKOnly, "Validation"
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 7) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")), "#,##0.00")
                    .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 13), ",", "")) * Val(.TextMatrix(i, 2)), "#,##0.00")
                    .TextMatrix(i, 6) = ""
                Next i
                txtRefNo.Text = ""
                lblRef.Caption = "Reference No."
            End If
        End With
    End If
    Call ComputeTotal
End Sub

Private Sub flxOrd_Click()
    Dim MyDesc As String, PrvW As Integer, CrnW As Integer, MaxW As Long, MyBO As Long, MySQ As Long
    
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
                If flxOrd.TextMatrix(ActRow, 4) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 1) & "|" & flxOrd.TextMatrix(ActRow, 2) & "|" & flxOrd.TextMatrix(ActRow, 3) & "|" & flxOrd.TextMatrix(ActRow, 4) & "|" & flxOrd.TextMatrix(ActRow, 5) & "|" & flxOrd.TextMatrix(ActRow, 6) & "|" & flxOrd.TextMatrix(ActRow, 7) & "|" & flxOrd.TextMatrix(ActRow, 8) & "|" & flxOrd.TextMatrix(ActRow, 9) & "|" & flxOrd.TextMatrix(ActRow, 10) & "|" & flxOrd.TextMatrix(ActRow, 11) & "|" & flxOrd.TextMatrix(ActRow, 12)
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
        Case 6  'Discount
            With txtDiscount
                
                If txtRefNo.Text = "WHOLESALE DISCOUNT" Then
                    If flxOrd.TextMatrix(ActRow, 6) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 7)
                    'Display discount
                    txtNoDiscount.Top = flxOrd.CellTop + flxOrd.Top + flxOrd.CellHeight
                    txtNoDiscount.Left = flxOrd.CellLeft + flxOrd.Left
                    txtNoDiscount.Text = "Max. " & flxOrd.TextMatrix(ActRow, 17) & IIf(Right$(flxOrd.TextMatrix(ActRow, 17), 1) <> "%", "%", "")
                    txtNoDiscount.Visible = True
                    .ZOrder
                    'Display Textbox
                    .Top = flxOrd.CellTop + flxOrd.Top
                    .Left = flxOrd.CellLeft + flxOrd.Left
                    .Width = flxOrd.CellWidth
                    If flxOrd.Text <> "" Then
                        .Text = IIf(Right$(flxOrd.Text, 1) <> "%", flxOrd.Text, Left$(flxOrd.Text, Len(flxOrd.Text) - 1))
                    Else
                    '                ------------ Key Accounts Discount ------------
                    If txtRefNo.Text = "KEY ACCOUNTS DISCOUNT" Then
                        If flxOrd.TextMatrix(ActRow, 6) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 7)
                        'Display discount
                        txtNoDiscount.Top = flxOrd.CellTop + flxOrd.Top + flxOrd.CellHeight
                        txtNoDiscount.Left = flxOrd.CellLeft + flxOrd.Left
                        txtNoDiscount.Text = "Max. " & flxOrd.TextMatrix(ActRow, 19) & IIf(Right$(flxOrd.TextMatrix(ActRow, 19), 1) <> "%", "%", "")
                        txtNoDiscount.Visible = True
                        .ZOrder
                        'Display Textbox
                        .Top = flxOrd.CellTop + flxOrd.Top
                        .Left = flxOrd.CellLeft + flxOrd.Left
                        .Width = flxOrd.CellWidth
                        If flxOrd.Text <> "" Then
                            .Text = IIf(Right$(flxOrd.Text, 1) <> "%", flxOrd.Text, Left$(flxOrd.Text, Len(flxOrd.Text) - 1))
                    Else
                            .Text = flxOrd.Text
                        End If
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ZOrder
                        .SetFocus
                    End If
                    End If
                End If

                If flxOrd.TextMatrix(ActRow, 4) <> "" And Val(flxOrd.TextMatrix(ActRow, 12)) > 0 Then
                    If flxOrd.TextMatrix(ActRow, 6) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 7)
                    'Display discount
                    txtNoDiscount.Top = flxOrd.CellTop + flxOrd.Top + flxOrd.CellHeight
                    txtNoDiscount.Left = flxOrd.CellLeft + flxOrd.Left
                    txtNoDiscount.Text = "Max. " & flxOrd.TextMatrix(ActRow, 12) & IIf(Right$(flxOrd.TextMatrix(ActRow, 12), 1) <> "%", "%", "")
                    txtNoDiscount.Visible = True
                    .ZOrder
                    'Display Textbox
                    .Top = flxOrd.CellTop + flxOrd.Top
                    .Left = flxOrd.CellLeft + flxOrd.Left
                    .Width = flxOrd.CellWidth
                    If flxOrd.Text <> "" Then
                        .Text = IIf(Right$(flxOrd.Text, 1) <> "%", flxOrd.Text, Left$(flxOrd.Text, Len(flxOrd.Text) - 1))
                    Else
                        .Text = flxOrd.Text
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                ElseIf flxOrd.TextMatrix(ActRow, 4) <> "" And Val(flxOrd.TextMatrix(ActRow, 12)) = 0 And txtRefNo.Text <> "WHOLESALE DISCOUNT" Then
                    txtNoDiscount.Top = flxOrd.CellTop + flxOrd.Top + 250
                    txtNoDiscount.Left = flxOrd.CellLeft + flxOrd.Left
                    txtNoDiscount.Text = "No Discount"
                    txtNoDiscount.Visible = True
                    .ZOrder
                Else
                    If flxOrd.TextMatrix(ActRow, 4) <> "" And Val(flxOrd.TextMatrix(ActRow, 12)) = 0 And txtRefNo.Text <> "KEY ACCOUNTS DISCOUNT" Then
                    txtNoDiscount.Top = flxOrd.CellTop + flxOrd.Top + 250
                    txtNoDiscount.Left = flxOrd.CellLeft + flxOrd.Left
                    txtNoDiscount.Text = "No Discount"
                    txtNoDiscount.Visible = True
                    .ZOrder
                End If
                End If
            End With
            
        Case 7         'Allow Administrators to Change the Price ... just temporary
            With txtPrice
                '9/17/2015 DO NOT ALLOW CHANGE IN PRICE IS ZERO DISCOUNT
                'If Val(flxOrd.TextMatrix(ActRow, 12)) = 0 Then Exit Sub
                If txtRefNo.Text = "WHOLESALE DISCOUNT" Then Exit Sub
                If flxOrd.TextMatrix(ActRow, 4) <> "" Then 'And user = "Administrator" Then
                    If flxOrd.TextMatrix(ActRow, 7) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 7)
                    .Top = flxOrd.CellTop + flxOrd.Top
                    .Left = flxOrd.CellLeft + flxOrd.Left
                    .Width = flxOrd.CellWidth
                    .Text = flxOrd.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End If
            End With
    End Select
End Sub
                             
Private Sub flxOrd_DblClick()
    Dim eMsg As String
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
    
    On Error GoTo LocalErr
    If ActCol = 0 And txtSPSorderID.Text <> "" Then
        If flxOrd.TextMatrix(ActRow, 0) = "" Then
            eMsg = "Updating Order Status [Item on Showroom]"
            flxOrd.Row = ActRow: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
                      
        Else
            eMsg = "Updating Order Status [Item not yet in Showroom]"
            flxOrd.TextMatrix(ActRow, 0) = ""
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
                       
        End If
    End If
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub


Private Sub flxOrd_EnterCell()
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
End Sub

Private Sub flxOrd_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then Call flxOrd_Click
End Sub
                                                  
Private Sub flxOrd_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim AltDown
    
    AltDown = (Shift And vbAltMask) > 0
    If flxOrd.TextMatrix(ActRow, 6) <> "" And txtRefNo.Text = "WHOLESALE DISCOUNT" And ActCol = 7 And AltDown And KeyCode = vbKeyW Then
        tmrQtyTip.Enabled = True
    End If
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
    If KeyCode = vbKeyF9 And Not JustScroll Then Call cmdCashier_Click 'Key Accounts
End Sub

Private Sub flxOrd_LeaveCell()
    txtNoDiscount.Visible = False
    txtAppModel.Visible = False
    txtSQtyBck.Visible = False
End Sub

Private Sub flxOrd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If JustScroll Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuStatus)
        End If
    Else
        If Button = vbRightButton Then
            Call PopupMenu(mnuRows)
        End If
    End If
End Sub

Private Sub flxParts_Click()
    Dim PctDisc As Single, myVat As Single, i As Integer
    Dim DiscPrice As Currency, TtlPrice As Currency, Qry As String
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
        flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
        flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
        flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
        flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
        flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
        flxOrd.TextMatrix(ActRow, 14) = .TextMatrix(.Row, 9) 'COST
        flxOrd.TextMatrix(ActRow, 15) = .TextMatrix(.Row, 10) 'Remarks
        flxOrd.TextMatrix(ActRow, 16) = .TextMatrix(.Row, 11) 'Wholesale Price
        flxOrd.TextMatrix(ActRow, 17) = .TextMatrix(.Row, 12) 'Wholesale Max Discount
'        flxOrd.TextMatrix(ActRow, 18) = .TextMatrix(.Row, 14) 'Key Accounts Price
'        flxOrd.TextMatrix(ActRow, 19) = .TextMatrix(.Row, 15) 'Key Accounts Discount
        
        If Val(.TextMatrix(.Row, 9)) = 0 Then
            flxOrd.TextMatrix(ActRow, 7) = "" 'Price
            flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
            Else
                If ApplyWholeSalePriceInRetail And Val(.TextMatrix(.Row, 11)) > 0 Then
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 11)), "###,##0.00") 'Dealer Price
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 11)) 'Wholesale
            Else
                If ApplyKeyAcctsPriceInRetail And Val(.TextMatrix(.Row, 14)) > 0 Then
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 14)), "###,##0.00") 'Key Accounts Price
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 14)) 'Key Accounts
            Else
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            End If
          
        End If
        'Qty and SRP >0
        If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
            TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(Replace(flxOrd.TextMatrix(ActRow, 7), ",", ""))
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
        End If
    End With
End Sub

Private Sub flxParts_EnterCell()
    Dim PctDisc As Single, myVat As Single, i As Integer
    Dim DiscPrice As Currency, TtlPrice As Currency, Qry As String
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
        flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
        flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
        flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
        flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
        flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
        flxOrd.TextMatrix(ActRow, 14) = .TextMatrix(.Row, 9) 'COST
        flxOrd.TextMatrix(ActRow, 15) = .TextMatrix(.Row, 10) 'Remarks
        flxOrd.TextMatrix(ActRow, 16) = .TextMatrix(.Row, 11) 'Wholesale Price
        flxOrd.TextMatrix(ActRow, 17) = .TextMatrix(.Row, 12) 'Wholesale Price
'        flxOrd.TextMatrix(ActRow, 18) = .TextMatrix(.Row, 14) 'Key Accounts Price
'        flxOrd.TextMatrix(ActRow, 19) = .TextMatrix(.Row, 15) 'Key Accounts Discount
        If Val(.TextMatrix(.Row, 9)) = 0 Then
            flxOrd.TextMatrix(ActRow, 7) = "" 'Price
            flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
        Else
            If ApplyWholeSalePriceInRetail And Val(.TextMatrix(.Row, 11)) > 0 Then
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 11)), "###,##0.00") 'Dealer Price
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 11)) 'Wholesale
            Else
            If ApplyKeyAcctsPriceInRetail And Val(.TextMatrix(.Row, 14)) > 0 Then
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 14)), "###,##0.00") 'KEYACCTS
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 14)) 'Key Accounts
            Else
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            End If
        End If
        End If
        'Qty and SRP >0
        If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
            TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(Replace(flxOrd.TextMatrix(ActRow, 7), ",", ""))
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
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
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

Function getFrontLiner() As String
    Dim rsFrontline As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT EMPLOYEES.FIRSTNAME, EMPLOYEES.EMPID From Employees " & _
          "WHERE (((EMPLOYEES.EMPID)=" & EmpId & "));"
    'mySQL
    Qry = "SELECT EMPLOYEES.FIRSTNAME, EMPLOYEES.EMPID From Employees " & _
          "WHERE EMPLOYEES.EMPID=" & UserID & ";"
    
    With rsFrontline
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            getFrontLiner = !FirstName
        Else
            getFrontLiner = ""
        End If
        .Close
    End With
End Function

Private Sub DisplayPendingOrders()
    Dim rsPending As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#)) ORDER BY SPSALESORDER.ORDERNO;"
    'mySQL
    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
          "WHERE SPSALESORDER.INVOICENO Is Null AND SPSALESORDER.EMPID = " & UserID & " AND SPSALESORDER.ORDERDATE = '" & Format(Date, "yyyy-mm-dd") & "' ORDER BY SPSALESORDER.ORDERNO;"
    
    With rsPending
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxPND.Cols = 0
            Do While Not .EOF
                flxPND.Cols = flxPND.Cols + 1
                flxPND.ColAlignment(flxPND.Cols - 1) = 4
                flxPND.TextMatrix(0, flxPND.Cols - 1) = !ORDERNO
                .MoveNext
            Loop
        Else
            flxPND.Cols = 0
        End If
        .Close
    End With
End Sub

Private Sub flxParts_LostFocus()
Dim i As Long, j As Long

    'Validate existing PartNo
    For i = 1 To flxOrd.Rows - 1
        If i <> ActRow Then
            If flxOrd.TextMatrix(i, 4) = txtPartNo.Text Then
                txtPartNo.Visible = False
                flxParts.Visible = False
                Ans = MsgBox("Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number")
                For j = 1 To flxOrd.Cols - 1
                    flxOrd.TextMatrix(ActRow, j) = ""
                Next j
                flxOrd.Row = i
                flxOrd.Col = 4
                flxOrd.SetFocus
                txtPartNo.Tag = "PN exist"
                Exit Sub
            End If
        End If
    Next i
End Sub

Private Sub flxPND_Click()
    Dim rsOrder As New ADODB.Recordset
    Dim rsOrdDtls As New ADODB.Recordset
    Dim Qry As String, myRem As String
    
    Call ClearEntries
    If UserType <> "Administrator" Then
        If flxPND.Cols = 0 Then Exit Sub
        If flxPND.TextMatrix(0, flxPND.Col) = "" Then Exit Sub
        
        'Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & flxPND.TextMatrix(0, flxPND.Col) & ") AND ((SPSALESORDER.ORDERDATE)=#" & Date & "#) AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ")) ORDER BY SPSALESORDER.SPSORDERID DESC;"
        'mySQL
        Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & flxPND.TextMatrix(0, flxPND.Col) & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(Date, "yyyy-mm-dd") & "') AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & UserID & ")) ORDER BY SPSALESORDER.SPSORDERID DESC;"
        
    Else
        'Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & Val(cboOrderNo.Text) & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#) AND ((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & txtEmpID.Text & ")) ORDER BY SPSALESORDER.SPSORDERID DESC;"
        'mySQL
        Qry = "SELECT SPSALESORDER.* From SPSALESORDER " & _
              "WHERE SPSALESORDER.ORDERNO=" & Val(cboOrderNo.Text) & " AND SPSALESORDER.ORDERDATE='" & Format(txtDate.Text, "yyyy-mm-dd") & "' AND ((SPSALESORDER.INVOICENO) Is Null) AND SPSALESORDER.EMPID=" & txtEmpID.Text & " ORDER BY SPSALESORDER.SPSORDERID DESC;"
        
    End If
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
                .MoveFirst
                txtOrderno.Text = !ORDERNO
                txtDate.Text = Format(!ORDERDATE, "mm/dd/yyyy")
                txtFrnt.Text = !Frontliner
                txtSPSorderID.Text = !SPSORDERID
                txtCustID.Text = IIf(IsNull(!CUSTOMERID), "", !CUSTOMERID)
                txtRefNo.Text = IIf(IsNull(!REFERENCE), "", !REFERENCE)
                If txtRefNo.Text = "WHOLESALE DISCOUNT" Then
                    lblRef.Caption = "NOTE:"
                    txtPctDisc.Text = Format(!PctDisc, "#,##0.00")
                Else
                    If txtRefNo.Text = "KEY ACCOUNTS DISCOUNT" Then
                    lblRef.Caption = "NOTE:"
                    txtPctDisc.Text = Format(!PctDisc, "#,##0.00")
                    End If
                End If
                
                txtInvNo.Text = IIf(IsNull(!InvoiceNO), "", !InvoiceNO)
                txtReqID.Text = IIf(IsNull(!ReqID), "", !ReqID)
                txtCusType.Text = IIf(IsNull(!Ctype), "", !Ctype)
                txtCustName.Text = IIf(IsNull(!CUSTOMER), "", !CUSTOMER)
                cboAddress.Text = IIf(IsNull(!address), "", !address)
                EMPID = rsOrder!EMPID
                txtSPSorderID.Text = !SPSORDERID
                
                    If Val(txtCustID.Text) > 0 Then
                        txtPayScheme.Text = IIf(IsNull(!PayScheme), "", !PayScheme)
                        txtBillAmt.Text = IIf(IsNull(!BILLAMT), "", Format(!BILLAMT, "#,##0.00"))
                        txtTerms.Text = IIf(IsNull(!Terms), "", !Terms)
                        txtNPays.Text = IIf(IsNull(!Npays), "", !Npays)
                        txtPayStart.Text = IIf(IsNull(!PayStart), "", Format(!PayStart, "mm/dd/yyyy"))
                        myRem = IIf(IsNull(!Remarks), "", !Remarks)
                        
                        If myRem = "CUSTOMER W/ DP" Then
                            lblRef.Caption = "O.R. Number"
                            lblCustomer.Caption = "Customer w/ DP"
                            lblRemarks.Caption = IIf(IsNull(!Remarks), "", !Remarks)
                            chkDP.Value = 1
                            Call GetCustomerTerm
                        Else
                            Call GetCustomerTerm
                            lblRef.Caption = "P.O. Number"
                            If txtCusType.Text = "WHOLESALE" Then
                                cmdRemoveWS.Visible = True
                                lblCustomer.Caption = "Wholesale Cust."
                                If Not IsNull(!Remarks) Then
                                    If !Remarks = "CASH" Then
                                        OptCash.Value = True
                                        OptOnAcct.Value = False
                                        ApplyWholeSalePriceInRetail = True
                                    Else
                                        OptCash.Value = False
                                        OptOnAcct.Value = True
                                        ApplyWholeSalePriceInRetail = True
                                        lblTerms.Caption = "Terms: " & txtTerms.Text & " Days @ " & Format(Val(txtBillAmt.Text), "#,##0.00") & " " & txtPayScheme.Text
                                    End If
                                End If
                            Else
                                lblCustomer.Caption = "On-Acct. Customer"
                                lblRemarks.Caption = IIf(IsNull(!Remarks), "", !Remarks)
                            End If
                        End If
                    Else
                        Call GetCustomerTerm
                        lblRef.Caption = "P.O. Number"
                        If txtCusType.Text = "KEY ACCOUNTS" Then
                            cmdRemoveWS.Visible = True
                            lblCustomer.Caption = "Key Accounts Cust."
                            If Not IsNull(!Remarks) Then
                                If !Remarks = "CASH" Then
                                    OptCash.Value = True
                                    OptOnAcct.Value = False
                                    ApplyKeyAcctsPriceInRetail = True
                                Else
                                    OptCash.Value = False
                                    OptOnAcct.Value = True
                                    ApplyKeyAcctsPriceInRetail = True
                                    lblTerms.Caption = "Terms: " & txtTerms.Text & " Days @ " & Format(Val(txtBillAmt.Text), "#,##0.00") & " " & txtPayScheme.Text
                                End If
                            End If
                        Else
                                lblCustomer.Caption = "On-Acct. Customer"
                                lblRemarks.Caption = IIf(IsNull(!Remarks), "", !Remarks)
                        End If
                    End If
                If Val(txtReqID.Text) > 0 Then
                     lblRef.Caption = "Requisition No."
                     txtRefNo.Text = GetReqInfo(txtReqID.Text, 2)
                     txtCustName.Text = GetReqInfo(txtReqID.Text, 1)
                Else
                     'txtCustName.Text = ""
                     'cboAddress.Text = ""
                     txtPayScheme.Text = ""
                     txtBillAmt.Text = ""
                     txtTerms.Text = ""
                     txtNPays.Text = ""
                     txtPayStart.Text = ""
                End If
                If txtOrderno.Text = "" Then
                   'txtDate.Text = Format(Date, "mm/dd/yyyy")
                   txtFrnt.Text = ""
                   txtSPSorderID.Text = ""
                   txtPayScheme.Text = ""
                   txtBillAmt.Text = ""
                   txtTerms.Text = ""
                   txtNPays.Text = ""
                   txtPayStart.Text = ""
                   'lblRemarks.Caption = ""
                   txtCustID.Text = ""
                   lblRef.Caption = ""
                   txtInvNo.Text = ""
                   txtReqID.Text = ""
                   txtPOno.Text = ""
                   lblRef.Caption = ""
                   'EmpId = 0
                   Call DisplayPendingOrders
                End If
         End If
        .Close
    End With
    If txtSPSorderID.Text = "" Then Exit Sub
    'Qry = "SELECT SPSORDERDETAIL.STATUS, SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.DISCOUNT, SPSORDERDETAIL.PRICE, [QTY]*[PRICE] AS TTLPRICE, SPAREPARTS.LOCATION, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.STOREPRICE, SPAREPARTS.COST " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    Qry = "SELECT SPSORDERDETAIL.STATUS, SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.DISCOUNT, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.QTY*SPSORDERDETAIL.PRICE AS TTLPRICE, SPAREPARTS.LOCATION, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.STOREPRICE, SPAREPARTS.COST, SPAREPARTS.REMARKS, SPAREPARTS.DLRPRICE, SPAREPARTS.WSMAXDISCOUNT, SPAREPARTS.KEYACCTS, SPAREPARTS.KAMAXDISCOUNT " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    
    With rsOrdDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrd.Rows = 1
            Do While Not .EOF
                flxOrd.AddItem IIf(IsNull(!STATUS), "", !STATUS) & vbTab & !QtyOnHand & vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(!DISCOUNT > 0, !DISCOUNT & "%", "") & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparepartID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MaxDiscount), "", !MaxDiscount) & vbTab & !STOREPRICE & vbTab & !Cost & vbTab & IIf(IsNull(!Remarks), "", !Remarks) & vbTab & IIf(IsNull(!DlrPrice), "", !DlrPrice) & vbTab & IIf(IsNull(!WSMAXDISCOUNT), "", !WSMAXDISCOUNT) & vbTab & IIf(IsNull(!KEYACCTS), "", !KEYACCTS) & vbTab & IIf(IsNull(!KAMAXDISCOUNT), "", !KAMAXDISCOUNT)
                If flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = "FOUND" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = "DROP" Then
                    flxOrd.Row = flxOrd.Rows - 1: flxOrd.Col = 0: flxOrd.CellFontName = "Wingdings": flxOrd.CellFontSize = 12: flxOrd.Text = lblChk.Caption
                Else
                    flxOrd.TextMatrix(flxOrd.Rows - 1, 0) = ""
                End If
                .MoveNext
            Loop
            Call ComputeTotal
        Else
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        End If
        .Close
    End With
    
    cmdCancel.Visible = True
    'cmdCashier.Visible = True
    cmdWareHouse.Visible = True
    cmdAdvFind.Visible = False
    cmdReq.Visible = False
    cmdPO.Visible = False
    cmdWholesale.Visible = False
End Sub

Function GetReqInfo(MyID As Long, MyCode As Integer) As String
    Dim rsReqCust As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT REQUISITION.CUSTOMER, REQUISITION.REQNO From REQUISITION WHERE (((REQUISITION.REQID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT REQUISITION.CUSTOMER, REQUISITION.REQNO From REQUISITION WHERE REQUISITION.REQID=" & MyID & ";"
    
    With rsReqCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If MyCode = 1 Then
                GetReqInfo = IIf(IsNull(!CUSTOMER), "", !CUSTOMER)
            Else
                GetReqInfo = IIf(IsNull(!REQNO), "", !REQNO)
            End If
        Else
            GetReqInfo = ""
        End If
        .Close
    End With
End Function

Private Sub GetCustomerTerm()
    Dim rsCust As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL From CUSTOMER " & _
          "WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL From CUSTOMER " & _
          "WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
    
    With rsCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtCustName.Text = !CustomerName
            cboAddress.Text = IIf(cboAddress.Text <> "", cboAddress.Text, !address)
            txtOSBal.Text = !OSbal
        End If
        .Close
    End With

    If txtTerms.Text = "" Then Call GetCustomerDefaultTerm
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
    If KeyCode = vbKeyF9 And Not JustScroll Then Call cmdCashier_Click 'Key Accounts
End Sub

Private Sub Form_Load()
    Dim rsPartsMan As New ADODB.Recordset
    Dim rsAdrs As New ADODB.Recordset
    Dim Qry As String
    
    ActRow = 0
    ActCol = 0
    JustScroll = True
    WSdiscountID = 0
    
    SPSorder.Caption = SysBranch & ": Spare Parts Frontline"
    lblToolTip.ToolTipText = MotoRep.Caption
    'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='VAT'));"
    'mySQL
    Qry = "SELECT settings.Value From settings " & _
          "Where settings.REMARKS = 'December Execom' AND settings.DESCRIPTION = 'Apply Wholesale@Retail'"
    If GetFieldValue(Qry) = "12/13/2016" Then
        ApplyWholeSalePriceInRetail = True
    Else
        ApplyWholeSalePriceInRetail = False
    End If
    'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='VAT';"
    txtVat.Text = 12    'GetFieldValue(Qry)
    'If rsParts.State = adStateClosed Then Timer1.Enabled = True
    If UserType = "Administrator" Then
        Call DisplayMainButton(2)
        flxPND.Visible = False
        'Shape1.Visible = False
        Shape2.Visible = False
        fraAdmin.Visible = True
        'Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER Where (((PCUSER.PartsMan) = True)) ORDER BY PCUSER.USERNAME;"
        'mySQL
        'Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER Where PCUSER.PartsMan = True ORDER BY PCUSER.USERNAME;"
        Qry = "SELECT DISTINCT pcuser.USERNAME, spsalesorder.EmpId From pcuser INNER JOIN spsalesorder ON (pcuser.EMPID = spsalesorder.EMPID) " & _
              "Where spsalesorder.ORDERDATE = '" & Format$(Date, "yyyy-mm-dd") & "' Order By pcuser.UserName"
        With rsPartsMan
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                cboPartsMan.Clear
                Do While Not .EOF
                    cboPartsMan.AddItem !UserName
                    cboPartsMan.ItemData(cboPartsMan.ListCount - 1) = !EMPID
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        Call DisplayPendingOrders
        If flxPND.Cols = 0 Then
            Call DisplayMainButton(1)
            txtOrderno.Text = "<Auto>"
            'flxOrd.SetFocus
            flxOrd.Col = 4
            AddNew = True
        Else
            Call DisplayMainButton(2)
        End If
    End If
    txtFrnt.Text = getFrontLiner
    'txtDate.Text = Format(Date, "mm/dd/yyyy") 'txtDate.Text = Format(Date, "mm/dd/yyyy")
    txtDate.Text = Format(ServerTime, "mm/dd/yyyy")
    Qry = "SELECT concat(settings.VALUE,', ', settings.REMARKS) as Address From settings Where settings.DESCRIPTION LIKE 'REGION%' Order By settings.Value"
    With rsAdrs
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboAddress.AddItem !address
                .MoveNext
            Loop
        End If
        .Close
    End With
    If SysBranchID = 14 Or SysBranchID = 26 Then
        chkClub.Visible = True
    End If
    
    'MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Order"
End Sub

Private Sub DisplayMainButton(MyCode As Integer)
    If MyCode = 1 Then
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        Call LockData(False)
        cmdPO.Visible = True
        cmdWholesale.Visible = True
        cmdReq.Visible = True
        cmdAdvFind.Visible = True
        cmdCreditCard.Visible = True
        cmdCancel.Visible = False
        cmdCashier.Visible = True
        cmdWareHouse.Visible = False
    Else
        cmdTrans(3).Visible = False
        cmdTrans(4).Visible = False
        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = True
        cmdTrans(2).Visible = True
        Call LockData(True)
        cmdPO.Visible = False
        cmdWholesale.Visible = False
        cmdReq.Visible = False
        cmdAdvFind.Visible = False
        cmdCancel.Visible = False
        cmdCashier.Visible = False
        cmdWareHouse.Visible = False
        cmdCreditCard.Visible = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ProperExit Then
        If rsParts.State = adStateOpen Then rsParts.Close
        ProperExit = False
    ElseIf JustScroll Then
        Call cmdTrans_Click(2)
        ProperExit = False
    Else
        Cancel = True
    End If
End Sub

Private Sub mnuAdd_Click()
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    Else
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
Private Sub SaveSalesOrder()
    Dim i As Integer, myTtl As Single, Ans As Integer, NewOnAcct As Boolean
    Dim eMsg As String, Qry As String, myDisc As Currency, MyUser As String
    Dim ReqID As Long, myRem As String, CustID As Long, MyRef As String, MyPSkm As String, MyPStart As String
    Dim MyBillAmt As Double, MyTerm As Integer, MyNpay As Integer, Pctd As Integer, OrdersNo As Long
                                  
    'Saving Table Spare Parts Invoice
    On Error GoTo LocalErr
    myTtl = Val(Replace(txtGrandTotal.Text, ",", ""))
    If InStr(txtPctDisc.Text, "%") > 0 Then
        Pctd = Val(Left$(txtPctDisc.Text, Len(txtPctDisc.Text) - 1))
    Else
        Pctd = Val(txtPctDisc.Text)
    End If
    If Val(txtReqID.Text) > 0 Then
        ReqID = txtReqID.Text
        myRem = lblRemarks.Caption
    Else
        ReqID = 0
        myRem = ""
    End If
    
    If AddNew Then
        eMsg = "Saving New Record for SP Sales Order"
        'old txtOrderno.Text = GetOrderNo + 1
        'txtOrderno.Text = GetOrdersNo
        'LYRcon.Execute "INSERT INTO SPSalesOrder (ORDERNO, ORDERDATE, TotalPrice, LessDiscPrice, GrossPrice, InDiscAmt, Ctype, PctDisc, Frontliner, EmpId, ReqID, Remarks) " & _
        "VALUES (" & Val(txtOrderno.Text) & ", #" & txtDate.Text & "#, " & myTtl & ", " & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", " & Val(Replace(txtTotal.Text, ",", "")) & ", " & Val(txtTtlDisc.Text) & ", '" & txtCusType.Text & "', " & Pctd & ", '" & txtFrnt.Text & "', " & EmpId & ", " & ReqID & ", '" & myRem & "');"
        
        'mySQL
        'Update Max order on previous date
        Call UpdateNumOrder
        myStation = ComputerName
        myStation = Left$(myStation, Len(myStation) - 1)
        
        If myCardNo <> "" Then
            LYRcon.Execute "INSERT INTO SPSalesOrder (ORDERDATE, TotalPrice, LessDiscPrice, GrossPrice, InDiscAmt, Ctype, PctDisc, Frontliner, EmpId, ReqID, Remarks, WorkStation, Customer, Address) " & _
            "VALUES ('" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & myTtl & ", " & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", " & Val(Replace(txtTotal.Text, ",", "")) & ", " & Val(txtTtlDisc.Text) & ", '" & txtCusType.Text & "', " & Pctd & ", '" & txtFrnt.Text & "', " & UserID & ", " & ReqID & ", '" & myCardPointsType & "', '" & myStation & "', '" & txtCustName.Text & "', '" & cboAddress.Text & "');"
        Else
            If txtRefNo.Text = "KEY ACCOUNTS DISCOUNT" Or txtRefNo.Text = "WHOLESALE DISCOUNT" Then
            LYRcon.Execute "INSERT INTO SPSalesOrder (ORDERDATE, TotalPrice, LessDiscPrice, GrossPrice, InDiscAmt, Ctype, PctDisc, Frontliner, EmpId, ReqID, Remarks, WorkStation, Customer, Address) " & _
            "VALUES ('" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & myTtl & ", " & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", " & Val(Replace(txtTotal.Text, ",", "")) & ", " & Val(txtTtlDisc.Text) & ", '" & txtCusType.Text & "', " & Pctd & ", '" & txtFrnt.Text & "', " & UserID & ", " & ReqID & ", '" & myRem & "', '" & myStation & "', '" & txtCustName.Text & "', '" & cboAddress.Text & "');"
        Else
            LYRcon.Execute "INSERT INTO SPSalesOrder (ORDERDATE, TotalPrice, LessDiscPrice, GrossPrice, InDiscAmt, Ctype, PctDisc, Frontliner, EmpId, ReqID, Remarks, WorkStation, Customer, Address) " & _
            "VALUES ('" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & myTtl & ", " & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", " & Val(Replace(txtTotal.Text, ",", "")) & ", " & Val(txtTtlDisc.Text) & ", '" & txtCusType.Text & "', " & Pctd & ", '" & txtFrnt.Text & "', " & UserID & ", " & ReqID & ", '" & myRem & "', '" & myStation & "', '" & txtCustName.Text & "', '" & cboAddress.Text & "');"
            End If
        End If
       'Generate Order number
       '-------------------------------------------------------------
        Qry = "SELECT MAX(spsalesorder.SPSORDERID) As CrtOrder " & _
                "From " & _
                    "spsalesorder " & _
                "Where " & _
                    "spsalesorder.ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' "
                'AND spsalesorder.WorkStation = '" & myStation & "';"

            MyCrtOrder = Val(GetFieldValue(Qry))
            Qry = "SELECT settings.`VALUE`,settings.DESCRIPTION,settings.REMARKS,settings.UPDATEDBY," & _
                "settings.LASTUPDATE From settings Where settings.DESCRIPTION = 'SALES ORDER';"
            OrdersNo = MyCrtOrder - Val(GetFieldValue(Qry))
            
        LYRcon.Execute "UPDATE spsalesorder SET ORDERNO = " & OrdersNo & " Where spsalesorder.SPSORDERID = " & MyCrtOrder & ";"
        '-------------------------------------------------------------
        'Old Qry = "SELECT SPSALESORDER.SPSORDERID From SPSALESORDER " & _
              "WHERE (SPSALESORDER.ORDERNO=" & txtOrderno.Text & " AND SPSALESORDER.ORDERDATE ='" & Format(txtDate.Text, "yyyy-mm-dd") & "' SPSALESORDER.WORKSTATION='" & ComputerName & "');"
        txtOrderno = OrdersNo
        txtSPSorderID.Text = MyCrtOrder
        
    Else
        eMsg = "Updating Record for SP Sales Order"
        MyUser = getFrontLiner
        If txtFrnt.Text <> MyUser Then
            If InStr(txtFrnt.Text, "/") > 0 Then
                MyUser = txtFrnt.Text
            Else
                MyUser = txtFrnt.Text & "/" & MyUser
            End If
        End If
        If myCardNo <> "" Then
            LYRcon.Execute "UPDATE SPSalesOrder SET ORDERNO = " & Val(txtOrderno.Text) & ", ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', TotalPrice = " & myTtl & ", LessDiscPrice = " & Val(Replace(txtPctDisc.Text, ",", "")) & ", GrossPrice = " & Val(Replace(txtTotal.Text, ",", "")) & ", InDiscAmt = " & Val(txtTtlDisc.Text) & ", Ctype =  '" & Trim(txtCusType.Text) & "', PctDisc = " & Pctd & ", FRONTLINER = '" & MyUser & "', " & _
            "CUSTOMER = '" & txtCustName.Text & "', ADDRESS = '" & cboAddress.Text & "' WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
        Else
            LYRcon.Execute "UPDATE SPSalesOrder SET ORDERNO = " & Val(txtOrderno.Text) & ", ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', TotalPrice = " & myTtl & ", LessDiscPrice = " & Val(Replace(txtPctDisc.Text, ",", "")) & ", GrossPrice = " & Val(Replace(txtTotal.Text, ",", "")) & ", InDiscAmt = " & Val(txtTtlDisc.Text) & ", Ctype = '" & txtCusType.Text & "', PctDisc = " & Pctd & ", FRONTLINER = '" & MyUser & "', " & _
            "CUSTOMER = '" & txtCustName.Text & "', ADDRESS = '" & cboAddress.Text & "' WHERE SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ";"
        End If
    End If
    If Not AddNew Then
         'Reset Records On-Account/Cash/Quick Service
         'Whole Sale Cash
        
        LYRcon.Execute "UPDATE SPSalesOrder SET  Remarks = '', CUSTOMERID = 0, PAYSCHEME = NULL, TERMS = 0, NPAYS = 0, PAYSTART = NULL, BILLAMT = 0 " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
    
        'Sales Order On-Account
        
        LYRcon.Execute "UPDATE SPSalesOrder SET  ReqID = 0, Remarks = '', CUSTOMERID = 0, Reference = '', PayScheme = 0, BillAmt = 0, Terms = 0, Npays = 0, PayStart = NULL " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"

    End If
    If myCardNo = "" Then
        If txtCusType.Text = "WHOLESALE" Or txtCusType.Text = "KEY ACCOUNTS" Then
            If OptCash.Value = True Then
                myRem = "CASH"
                CustID = txtCustID.Text
                MyRef = txtRefNo.Text
            Else
                myRem = "ONACCOUNT"
                CustID = Val(txtCustID.Text)
                MyRef = txtRefNo.Text
                MyPSkm = txtPayScheme.Text
                If txtNPays.Text = "" Then
                    MyBillAmt = myTtl
                Else
                    MyBillAmt = myTtl / Val(txtNPays.Text)
                End If
                MyTerm = Val(txtTerms.Text)
                MyNpay = Val(txtNPays.Text)
                If IsDate(txtPayStart.Text) Then
                    MyPStart = txtPayStart.Text
                Else
                    MyPStart = DateAdd("m", 1, DateValue(txtDate.Text))
                End If
            End If
        ElseIf Val(txtCustID.Text) > 0 Then
            CustID = txtCustID.Text
            MyPSkm = txtPayScheme.Text
            If Val(txtBillAmt.Text) > 0 Then
                MyBillAmt = Val(Replace(txtBillAmt.Text, ",", ""))
            Else
                MyBillAmt = myTtl / Val(txtNPays.Text)
            End If
            MyTerm = Val(txtTerms.Text)
            MyNpay = Val(txtNPays.Text)
            If IsDate(txtPayStart.Text) Then
                MyPStart = txtPayStart.Text
            Else
                MyPStart = DateAdd("m", 1, DateValue(txtDate.Text))
            End If
            MyRef = txtRefNo.Text
            If lblRef.Caption = "CARD Number" Then
                myRem = myCardPointsType
            End If
            If chkDP.Value = 1 Then
                myRem = "CUSTOMER W/ DP"
            End If
        End If
    Else
        If OptCash.Value = True Then
            myRem = myCardPointsType
            CustID = myCardCustomerID
            MyRef = myCardNo
        End If
    End If
    If Val(txtCustID.Text) > 0 And OptCash.Value = True Then
        eMsg = "Updating Record for SP Sales Order Wholesale Cash"
        LYRcon.Execute "UPDATE SPSalesOrder SET  Remarks = '" & myRem & "', CUSTOMERID = " & CustID & ", Reference = '" & MyRef & "', PAYSCHEME = NULL, TERMS = 0, NPAYS = 0, PAYSTART = NULL, BILLAMT = 0 " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
    ElseIf Val(txtCustID) > 0 And myCardNo <> "" Then
        eMsg = "Updating Record for SP Sales Order On Promo Card Customer"
        LYRcon.Execute "UPDATE SPSalesOrder SET CUSTOMERID = " & Val(myCardCustomerID) & ", Reference = '" & myCardNo & "', Remarks = '" & myCardPointsType & "' " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
    ElseIf Val(txtCustID.Text) > 0 Then
        eMsg = "Updating Record for SP Sales Order On Account Customer"
        'LYRcon.Execute "UPDATE SPSalesOrder SET  ReqID = " & ReqID & ", Remarks = '" & myRem & "', CUSTOMERID = " & CustID & ", Reference = '" & MyRef & "', PayScheme = '" & MyPSkm & "', BillAmt = " & MyBillAmt & ", Terms = " & MyTerm & ", Npays = " & MyNpay & ", PayStart =" & IIf(IsDate(MyPStart), "#" & MyPStart & "#", Null) & " " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
        'mySQL
        LYRcon.Execute "UPDATE SPSalesOrder SET  ReqID = " & ReqID & ", Remarks = '" & myRem & "', CUSTOMERID = " & CustID & ", Reference = '" & MyRef & "', PayScheme = '" & MyPSkm & "', BillAmt = " & MyBillAmt & ", Terms = " & MyTerm & ", Npays = " & MyNpay & ", PayStart =" & IIf(IsDate(MyPStart), "'" & Format(MyPStart, "yyyy-mm-dd") & "'", Null) & " " & _
        "WHERE (SPSALESORDER.SPSORDERID=" & txtSPSorderID.Text & ");"
    End If
    
    
    'Saving Table Spare Parts Order Details
    If Not AddNew Then
        'LYRcon.Execute "DELETE SPSORDERDETAIL.* From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
        'mySQL
         LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                       "WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & ";"
     
    ElseIf Val(txtReqID.Text) > 0 Then
        'LYRcon.Execute "DELETE SPSORDERDETAIL.* From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.REQID)=" & Val(txtReqID.Text) & "));"
        'mySQL
        LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                       "WHERE SPSORDERDETAIL.REQID=" & Val(txtReqID.Text) & ";"
    
    End If
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
                If InStr(.TextMatrix(i, 6), "%") > 0 Then
                    myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                Else
                    myDisc = Val(.TextMatrix(i, 6))
                End If
                eMsg = "Line 1136: Adding New Sales Order Details @ row " & i
                LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT, REQID) " & _
                               "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & myDisc & ", " & Val(txtReqID.Text) & ");"
        Next i
    End With
    
    Exit Sub
LocalErr:
    If eMsg = "Saving SP Sales Order" Or Err.Number = 3219 Then
        Call SaveFromError3219
    ElseIf InStr(eMsg, "Line 1136") > 0 Then
        Call SaveDetailsOnly
    Else
        MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub
Private Sub SaveDetailsOnly()
    Dim i As Integer, myDisc As Currency
    On Error GoTo LocalErr
    'LYRcon.Execute "DELETE SPSORDERDETAIL.*, SPSORDERDETAIL.SPSORDERID From SPSORDERDETAIL " & _
                   "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & "));"
    'mySQL
    LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                   "WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & ";"
    
    With flxOrd
        For i = 1 To .Rows - 1
            If InStr(.TextMatrix(i, 6), "%") > 0 Then
                myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
            Else
                myDisc = Val(.TextMatrix(i, 6))
            End If
            LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT, REQID) " & _
                           "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & myDisc & ", " & Val(txtReqID.Text) & ");"
        Next i
    End With
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub
Private Sub SaveFromError3219()
    Dim rsOrder As New ADODB.Recordset
    Dim rsODetails As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, myTtl As Single, Ans As Integer, NewOnAcct As Boolean
    Dim eMsg As String, eOrderID As Long
                                  
    'Saving Table Spare Parts Invoice
    On Error GoTo LocalErr
    myTtl = Val(Replace(txtGrandTotal.Text, ",", ""))
    If Not AddNew Then
        AddNew = True
        eMsg = "Deleting Previous SP Sales Order ID " & txtSPSorderID.Text
        eOrderID = Val(txtSPSorderID.Text)
        'LYRcon.Execute "DELETE SPSALESORDER.* From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
        'mySQL
        LYRcon.Execute "DELETE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID=" & eOrderID & ";"
        
    End If
    eMsg = "Saving SP Sales Order RD2"
    
    'rsOrder.Open "SPSalesOrder", LYRcon, adOpenKeyset, adLockOptimistic
    'rsOrder.AddNew
    'txtOrderno.Text = GetOrderNo + 1
    'rsOrder!ORDERNO = Val(txtOrderno.Text)
    'rsOrder!ORDERDATE = txtDate.Text
    'rsOrder!TotalPrice = myTtl
    'rsOrder!LessDiscPrice = Val(Replace(txtLessDiscAmt.Text, ",", ""))
    'rsOrder!GrossPrice = Val(Replace(txtTotal.Text, ",", ""))
    'rsOrder!InDiscAmt = Val(txtTtlDisc.Text)
    'If InStr(txtPctDisc.Text, "%") > 0 Then
    '    rsOrder!PctDisc = Val(Left$(txtPctDisc.Text, Len(txtPctDisc.Text) - 1))
    'Else
    '    rsOrder!PctDisc = Val(txtPctDisc.Text)
    'End If
    'rsOrder!Frontliner = txtFrnt.Text
    'rsOrder!EmpId = EmpId
    'If Val(txtReqID.Text) > 0 Then
    '    rsOrder!ReqID = txtReqID.Text
    '    rsOrder!Remarks = lblRemarks.Caption
    'End If
    'If Val(txtCustID.Text) > 0 Then
    '    rsOrder!CUSTOMERID = txtCustID.Text
    '    rsOrder!PayScheme = txtPayScheme.Text
    '    rsOrder!BillAmt = Val(Replace(txtBillAmt.Text, ",", ""))
    '    rsOrder!Terms = Val(txtTerms.Text)
    '    rsOrder!Npays = Val(txtNPays.Text)
    '    If IsDate(txtPayStart.Text) Then rsOrder!PayStart = txtPayStart.Text
    '    rsOrder!Reference = txtRefNo.Text
    '    If chkDP.Value = 1 Then rsOrder!Remarks = "CUSTOMER W/ DP"
    'ElseIf Val(txtCustID.Text) = 0 And Not AddNew Then
    '    rsOrder!CUSTOMERID = 0
    '    rsOrder!PayScheme = ""
    '    rsOrder!Terms = 0
    '    rsOrder!Npays = 0
    '    rsOrder!Reference = ""
    '    rsOrder!Remarks = ""
    'End If
    'rsOrder.Update
    'txtSPSorderID.Text = rsOrder!SPSORDERID
    'rsOrder.Close
    
    'mySQl
    txtOrderno.Text = GetOrderNo + 1
    LYRcon.Execute "INSERT INTO spsalesorder( ORDERNO, ORDERDATE, TOTALPRICE, LESSDISCPRICE, GROSSPRICE, INDISCAMT)" & _
        "VALUES( " & Val(txtOrderno.Text) & ", '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & myTtl & ", " & _
        "" & Val(Replace(txtLessDiscAmt.Text, ",", "")) & ", " & Val(Replace(txtTotal.Text, ",", "")) & ", " & _
        "" & Val(txtTtlDisc.Text) & ");"
        
    If InStr(txtPctDisc.Text, "%") > 0 Then
        SPPctDisc = Val(Left$(txtPctDisc.Text, Len(txtPctDisc.Text) - 1))
    Else
        SPPctDisc = Val(txtPctDisc.Text)
    End If
    LYRcon.Execute "INSERT INTO spsalesorder( PCTDISC, FRONTLINER, EMPID)VALUES( " & SPPctDisc & ", '" & txtFrnt.Text & "', " & UserID & ");"
    If Val(txtReqID.Text) > 0 Then
        LYRcon.Execute "INSERT INTO spsalesorder( REQID, REMARKS )VALUES( " & txtReqID.Text & ", '" & lblRemarks.Caption & "');"
    End If
    If Val(txtCustID.Text) > 0 Then
        
        LYRcon.Execute "INSERT INTO spsalesorder( CUSTOMERID, PAYSCHEME, BILLAMT, TERMS, NPAYS, REFERENCE)" & _
            "VALUES( " & txtCustID.Text & ", '" & txtPayScheme.Text & "', " & Val(Replace(txtBillAmt.Text, ",", "")) & ", " & _
            "" & Val(txtTerms.Text) & ", " & Val(txtNPays.Text) & ", '" & txtRefNo.Text & "');"
        
        If IsDate(txtPayStart.Text) Then LYRcon.Execute "INSERT INTO spsalesorder(PAYSTART)VALUES('" & Format(PayStart, "yyyy-mm-dd") & "');"
        If chkDP.Value = 1 Then LYRcon.Execute "INSERT INTO spsalesorder(REMARKS)VALUES('CUSTOMER W/ DP');"
    
    ElseIf Val(txtCustID.Text) = 0 And Not AddNew Then
        
        LYRcon.Execute "INSERT INTO spsalesorder(CUSTOMERID, PAYSCHEME, TERMS, NPAYS, REFERENCE, REMARKS)VALUES(0, '', 0, 0, '', '');"
    
    End If
    Qry = "SELECT spsalesorder.SPSORDERID From spsalesorder Where spsalesorder.ORDERNO = " & Val(txtOrderno.Text) & " AND " & _
        "spsalesorder.ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "';"
    txtSPSorderID.Text = GetFieldValue(Qry)
    
    
    'Saving Table Spare Parts Order Details
    If Not AddNew Then
        'LYRcon.Execute "DELETE SPSORDERDETAIL.*, SPSORDERDETAIL.SPSORDERID From SPSORDERDETAIL " & _
                       "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & eOrderID & "));"
        'mySQL
        LYRcon.Execute "DELETE From SPSORDERDETAIL " & _
                       "WHERE SPSORDERDETAIL.SPSORDERID=" & eOrderID & ";"
        
    End If
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
            If Val(txtReqID.Text) <> 0 Then
                'Qry = "SELECT SPSORDERDETAIL.SPSORDERID, SPSORDERDETAIL.SPAREPARTID From SPSORDERDETAIL " & _
                      "WHERE (((SPSORDERDETAIL.REQID)=" & Val(txtReqID.Text) & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & .TextMatrix(i, 10) & "));"
                'mySQL
                Qry = "SELECT SPSORDERDETAIL.SPSORDERID, SPSORDERDETAIL.SPAREPARTID From SPSORDERDETAIL " & _
                      "WHERE (((SPSORDERDETAIL.REQID)=" & Val(txtReqID.Text) & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & .TextMatrix(i, 10) & "));"
                              
                If CtrlNumExist(Qry) Then
                    eMsg = "Line 1110: Updating Sales Order Details"
                    LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.SPSORDERID = " & txtSPSorderID.Text & ", SPSORDERDETAIL.QTY = " & .TextMatrix(i, 2) & ", SPSORDERDETAIL.PRICE = " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", SPSORDERDETAIL.STATUS = '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "' " & _
                                   "WHERE SPSORDERDETAIL.SPAREPARTID=" & .TextMatrix(i, 10) & " AND SPSORDERDETAIL.REQID=" & txtReqID.Text & ";"
                Else
                    If InStr(.TextMatrix(i, 6), "%") > 0 Then
                        myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                    Else
                        myDisc = Val(.TextMatrix(i, 6))
                    End If
                    eMsg = "Line 1119: Adding record to Sales Order Details"
                    LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, REQID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                                   "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & Val(txtReqID.Text) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & myDisc & ");"
                End If
            Else
                If InStr(.TextMatrix(i, 6), "%") > 0 Then
                    myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
                Else
                    myDisc = Val(.TextMatrix(i, 6))
                End If
                eMsg = "Line 1136: Adding New Sales Order Details"
                LYRcon.Execute "INSERT INTO SPSORDERDETAIL ( SPSORDERID, SPAREPARTID, QTY, UNIT, PRICE, STATUS, DISCOUNT) " & _
                               "VALUES (" & txtSPSorderID.Text & ", " & .TextMatrix(i, 10) & ", " & .TextMatrix(i, 2) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & IIf(.TextMatrix(i, 0) <> "", "FOUND", "BLANK") & "', " & myDisc & ");"
            End If
        Next i
    End With
    
    Qry = "SELECT SPSALESORDER.SPSORDERID From SPSALESORDER WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
    If CtrlNumExist(Qry) Then
        'LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                       "WHERE (((SPSALESORDER.SPSORDERID)=" & eOrderID & "));"
        'mySQL
        LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = 'CANCELLED' " & _
                       "WHERE SPSALESORDER.SPSORDERID=" & eOrderID & ";"
        
    End If
    MsgBox "This is Order Number " & txtOrderno.Text, vbInformation + vbOKOnly, "New Order was made"
    Exit Sub
LocalErr:
    If InStr(eMsg, "Deleting Previous SP Sales Order ID") > 0 Then
        Resume Next
    Else
        MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Function GetOrderNo() As Integer
    Dim rsMyOrder As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT Max(SPSALESORDER.ORDERNO) AS MaxOfORDERNO From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#));"
    'mySQL
    Qry = "SELECT Max(SPSALESORDER.ORDERNO) AS MaxOfORDERNO From SPSALESORDER " & _
          "WHERE SPSALESORDER.ORDERDATE='" & Format(txtDate.Text, "yyyy-mm-dd") & "';"
    
    With rsMyOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetOrderNo = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            GetOrderNo = 0
        End If
        .Close
    End With
End Function

Public Sub CreateTempTableOrder()
On Error GoTo er
        LYRcon.Execute "Drop  Table IF  Exists sp_mrcretails.TMP_SPID;"

         DoEvents
       LYRcon.Execute " Create  Table IF Not Exists sp_mrcretails.TMP_SPID( " & _
        "   SPID Int(11),              " & _
        "   QTY Int(11)     " & _
        ");"

      ' "   PMATRIXID Int(11),           "
        Qry = "insert into sp_mrcretails.TMP_SPID(SPID , QTY) VALUES "
        
        For i = 1 To flxOrd.Rows - 1
            If i > 1 Then Qry = Qry & ","
            Qry = Qry & "('" & Val(flxOrd.TextMatrix(i, 10)) & "', '" & Val(flxOrd.TextMatrix(i, 2)) & "')"
        Next
        
       Call LYRcon.Execute(Qry)

Exit Sub
er:
 Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub
Function ValidateSalesOrder() As String
    Dim MyErr As String, Qry As String, myPN As String
    Dim addnewvalue As Integer, rsValidate As New ADODB.Recordset
    Dim i As Long, j As Long
    
    MyErr = ""
    myPN = ""
    addnewvalue = 0
    If myCardNo = "" Then
        If chkDP.Value = 0 Then
            If txtRefNo.Text <> "WHOLESALE DISCOUNT" Or txtRefNo.Text <> "KEY ACCOUNTS DISCOUNT" Then
                If Val(txtCustID.Text) > 0 And chkWholesaleCash.Value = 0 And txtPayScheme.Text = "" Then MyErr = "No Terms specified for On-Account sales order."
            End If
        End If
    End If
    
    For i = 1 To flxOrd.Rows - 1
        For j = 1 To flxOrd.Rows - 1
            If i <> j Then
                If flxOrd.TextMatrix(i, 4) = flxOrd.TextMatrix(j, 4) Then
                    'MsgBox "Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number"
                    MyErr = MyErr & vbCrLf & "Part Number " & flxOrd.TextMatrix(i, 4) & " already exist in row " & i
                End If
            End If
        Next j
    Next i
    
    If Not IsDate(txtDate.Text) Then MyErr = MyErr & vbCrLf & "Invalid entry on field Date"
    If txtFrnt.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Frontliner"
    If txtCustName.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Customer Name"
    If cboAddress.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Address"
    
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" And flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then
        If flxOrd.Rows = 2 Then
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        Else
            flxOrd.RemoveItem flxOrd.Rows - 1
        End If
    End If
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row of Ordered Spare Parts"
    If AddNew = True Then
        addnewvalue = 1
    Else
        addnewvalue = 0
    End If
    For i = 1 To flxOrd.Rows - 1
        Call CreateTempTableOrder
        Qry = "call sp_mrcretails.sp_CheckQty('" & sysBranchCode & "','" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & Val(addnewvalue) & "')"
        
        If rsValidate.State = adStateOpen Then rsValidate.Close
        With rsValidate
            .Open Qry, LYRcon, adOpenKeyset, adLockReadOnly
            
            If .RecordCount > 0 Then
                If !Remarks <> "" Then
                    MyErr = MyErr & vbCrLf & "Qty must not Exceeded in Qty On-Hand of the Following Spare Parts Sales: '" & !Remarks & "' "
                End If
            End If
            If .State <> 0 Then .Close
        End With
        If Val(flxOrd.TextMatrix(i, 2)) = 0 Then MyErr = MyErr & vbCrLf & "Ordered Spare Parts Sales @ Row " & i & " no Qty.!"
        'If Val(flxOrd.TextMatrix(i, 7)) = 0 Then MyErr = MyErr & vbCrLf & "Ordered Spare Parts Sales @ Row " & i & " no Price!"
    Next i
    If ApplyWholeSalePriceInRetail = True And AddNew = False Then
        Call ComputeWholesaleDiscount
    Else
        If ApplyKeyAcctsPriceInRetail = True And AddNew = False Then 'KEY ACCTS
            Call ComputeWholesaleDiscount
        End If
    End If
    
    ValidateSalesOrder = MyErr
End Function

Private Sub LoadSpareParts()
    Dim rsSet As New ADODB.Recordset
    Dim Qry As String
    Dim pPct As Single
        

    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST, SPAREPARTS.WSMAXDISCOUNT, SPAREPARTS.KEYACCTS, SPAREPARTS.KAMAXDISCOUNT " & _
          "From SPAREPARTS WHERE SPAREPARTS.DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY SPAREPARTS.PARTNUMBER;"
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
End Sub


Private Sub RXResizer1_ResizeComplete()
    
    On Error GoTo LocalErr
    With flxOrd
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.05
        .ColWidth(2) = (.Width - 300) * 0.08
        .ColWidth(3) = (.Width - 300) * 0.08
        .ColWidth(4) = (.Width - 300) * 0.16
        .ColWidth(5) = (.Width - 300) * 0.2
        .ColWidth(6) = (.Width - 300) * 0.1
        .ColWidth(7) = (.Width - 300) * 0.1
        .ColWidth(8) = (.Width - 300) * 0.1
        .ColWidth(9) = (.Width - 300) * 0.12
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .Col = 6: .Row = 1
        txtTotal.Left = flxOrd.Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7)
        txtTotal.Width = .ColWidth(8)
        flxParts.ColWidth(2) = .ColWidth(4)
        Label3.Left = txtTotal.Left - (Label3.Width + 50)
        Label3.Top = txtTotal.Top
        Label3.FontSize = txtTotal.FontSize
    End With
    flxGridRef.ColWidth(0) = flxGridRef.Width * 0.25
    flxGridRef.ColWidth(1) = flxGridRef.Width * 0.4
    flxGridRef.ColWidth(2) = flxGridRef.Width * 0.35
    
    txtLessDiscAmt.Left = txtTotal.Left
    txtLessDiscAmt.Width = txtTotal.Width
    txtLessDiscAmt.Top = txtTotal.Top + txtTotal.Height
    txtGrandTotal.Left = txtTotal.Left
    txtGrandTotal.Width = txtTotal.Width
    txtGrandTotal.Top = txtLessDiscAmt.Top + txtLessDiscAmt.Height
    lblDiscount.Left = txtTotal.Left - (lblDiscount.Width + 50)
    lblDiscount.Top = txtLessDiscAmt.Top
    txtPctDisc.Left = lblDiscount.Left - (txtPctDisc.Width + 50)
    txtPctDisc.Top = lblDiscount.Top
    lblLess.Left = txtPctDisc.Left - (lblLess.Width + 50)
    lblLess.Top = txtPctDisc.Top
    lblGrandTtl.Left = txtGrandTotal.Left - (lblGrandTtl.Width + 50)
    lblGrandTtl.Top = txtGrandTotal.Top
    
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
    flxPND.Left = cmdTrans(2).Left + cmdTrans(2).Width + 1000 'Me.ScaleWidth - flxPND.Width
    flxPND.Width = Me.ScaleWidth - flxPND.Left
    Shape2.Left = flxPND.Left
    Shape2.Width = flxPND.Width
    'Shape1.Left = Shape2.Left - 240
    fraAdmin.Left = Frame1.Left + (Frame1.Width - fraAdmin.Width)
    Exit Sub
LocalErr:
    'Unload Me
End Sub



Private Sub tabPromoCard_Click(PreviousTab As Integer)
If tabPromoCard.Tab = 1 Then
    txtRPCardNo.Text = ""
    txtRPCustomer.Text = ""
    cboRPPointsType.Text = ""
    txtRPCardPoints.Text = ""
    myCardNo = ""
    myCardCustomerID = 0
Else
    txtPCCNo.Text = ""
    txtPCCustomer.Text = ""
    txtPCCPoints.Text = ""
    cboPCPointsType.Text = ""
    myCardNo = ""
    myCardCustomerID = 0
End If
End Sub

Private Sub Timer1_Timer()


    Dim Qry As String

    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST, SPAREPARTS.REMARKS, SPAREPARTS.DLRPRICE, SPAREPARTS.WSMAXDISCOUNT, SPAREPARTS.KEYACCTS, SPAREPARTS.KAMAXDISCOUNT " & _
          "From SPAREPARTS WHERE SPAREPARTS.DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    flxParts.Refresh
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
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        
    End With
    
    Timer1.Enabled = False

End Sub

Private Sub tmrQtyTip_Timer()
    Dim MyCost As Currency, PctDisc As Currency, i As Integer
    
    If tmrCount = 0 Then
        With txtQtyTip
            .Top = flxOrd.CellTop + flxOrd.Top + 290
            .Left = flxOrd.Left + flxOrd.ColWidth(0) + flxOrd.ColWidth(1) + flxOrd.ColWidth(2) + flxOrd.ColWidth(3) + flxOrd.ColWidth(4) + flxOrd.ColWidth(5) + flxOrd.ColWidth(6)
            MyCost = 0
            For i = 1 To flxOrd.Rows - 1    'Get Total Amount w/out Discount
                MyCost = MyCost + Val(flxOrd.TextMatrix(i, 13))
            Next i
            PctDisc = IIf(MyCost >= 10000, 1.03, 1.05)
            .Text = "Discount Amount: " & Format(Val(flxOrd.TextMatrix(ActRow, 13)) - Val(flxOrd.TextMatrix(ActRow, 16)), "#,##0.00")
'            .Text = "Wholesale Price " & Format(Val(flxOrd.TextMatrix(ActRow, 16)), "#,##0.00") & " * " & PctDisc & " = " & flxOrd.TextMatrix(ActRow, 7)
'            If Val(txtVat.Text) = 0 Then
'                .Text = "No VAT available"
'            Else
'                MyCost = Val(flxOrd.TextMatrix(ActRow, 14))
'                CostVat = MyCost * (1 + (Val(txtVat.Text) / 100))
'                .Text = "Cost with Vat: " & Format(CostVat, "#,##0.00")
'            End If
            .Visible = True
            .ZOrder
        End With
        tmrCount = 1
        tmrQtyTip.Interval = 5000
    Else
        txtQtyTip.Visible = False
        tmrQtyTip.Enabled = False
        tmrCount = 0
        tmrQtyTip.Interval = 300
    End If
End Sub

Private Sub txtBankCharge_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtBankCharge, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then cmdCCOk_Click
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
    If KeyCode = vbKeyF9 And Not JustScroll Then Call cmdCashier_Click 'Key Accounts
End Sub

Private Sub txtDiscount_GotFocus()
'      If SysBranchID = 14 And flxOrd.TextMatrix(ActRow, 15) = "YAMAHA GENUINE" Then
'            Ans = MsgBox("Do You Want to avail Yamaha Club Member Discount Price?", vbYesNo, "Yamaha Club Member")
'                If Ans = vbYes Then
'                    'If chkClub.Value = 0 Then chkClub.Value = 1
'                    flxOrd.TextMatrix(ActRow, 12) = 10
'                    txtDiscount.Tag = "Xcd100"
'                    txtDiscount.Text = flxOrd.TextMatrix(ActRow, 12)
'                End If
'    End If
    
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
        If txtDiscount.Text <> "" And txtDiscount.Tag <> txtDiscount.Text Then
            flxOrd.SetFocus
            flxOrd.Row = ActRow
            flxOrd.Col = 6
        ElseIf flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
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
        If Not ((SysBranchID = 14 Or SysBranchID = 26) And chkClub.Value = 1 And flxOrd.TextMatrix(ActRow, 15) = "YAMAHA GENUINE") Then
            
            If txtRefNo.Text = "WHOLESALE DISCOUNT" Then
                If Val(txtDiscount.Text) > Val(flxOrd.TextMatrix(ActRow, 17)) Then
                    MsgBox "Discount must not exceed " & flxOrd.TextMatrix(ActRow, 17) & "%", vbCritical + vbOKOnly, "Validation"
                    txtDiscount.Tag = "Xcd100"
                    txtDiscount.Text = flxOrd.TextMatrix(ActRow, 17)
                    Exit Sub
                End If
            Else
                If Val(txtDiscount.Text) > Val(flxOrd.TextMatrix(ActRow, 12)) Then
                    MsgBox "Discount must not exceed " & flxOrd.TextMatrix(ActRow, 12) & "%", vbCritical + vbOKOnly, "Validation"
                    txtDiscount.Tag = "Xcd100"
                    txtDiscount.Text = flxOrd.TextMatrix(ActRow, 12)
                    Exit Sub
                End If
            End If
        Else
            If Val(txtDiscount.Text) > Val(txtYClub.Text) Then
                MsgBox "Yamaha Club Member Discount must not exceed " & txtYClub.Text & "%", vbCritical + vbOKOnly, "Validation"
                txtDiscount.Tag = "Xcd100"
                txtDiscount.Text = flxOrd.TextMatrix(ActRow, 12)
                Exit Sub
            End If
        End If
            
        
        If Val(txtDiscount.Text) > 0 Then
            DiscAmt = Val(Replace(.TextMatrix(ActRow, 13), ",", "")) * (Val(txtDiscount.Text) / 100)
            DiscPrice = Val(Replace(.TextMatrix(ActRow, 13), ",", "")) - DiscAmt 'Discounted
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

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
    If KeyCode = vbKeyF9 And Not JustScroll Then Call cmdCashier_Click 'Key Accounts
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
        txtOrderno.SetFocus
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
    'if ActCol = 1
    If ActCol = 1 Or ActCol = 2 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        Call txtFlx_LostFocus
        If txtFlx.Tag = "XcdQty" Then txtFlx.Tag = "": Exit Sub
        'If all fields are filled then addrow focus on next entry
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            If ActRow = flxOrd.Rows - 1 Then flxOrd.AddItem ""
            flxOrd.Row = flxOrd.Rows - 1
            flxOrd.Col = 4
            flxOrd.SetFocus
        ElseIf flxOrd.TextMatrix(ActRow, 5) = "" Then   'No Description focus on Partnumber of current row
            flxOrd.Col = 4
            flxOrd.SetFocus
        ElseIf flxOrd.TextMatrix(ActRow, 2) = "" Then   'No Qty Focus on same column
            flxOrd.Col = 2
            flxOrd.SetFocus
        Else
            flxOrd.SetFocus
        End If
    End If
End Sub
Private Sub txtFlx_LostFocus()
    Dim AvlQty As Long
    
    If txtFlx.Tag = "skiplostfocus" Then
        txtFlx.Tag = ""
        Exit Sub
    End If
    With flxOrd
        Call TOT_PENDINGORDER(Format(txtDate.Text, "yyyy-mm-dd"), Val(flxOrd.TextMatrix(ActRow, 10)))
        AvlQty = Val(.TextMatrix(ActRow, 1)) - Val(txtTotPen.Text)
        If Val(txtFlx.Text) > AvlQty And .TextMatrix(ActRow, 10) <> "" Then
            MsgBox "Order Qty. must not exceed Qty. On-Hand (" & AvlQty & ")!", vbCritical + vbOKOnly, "Validation"
            txtFlx.Tag = "XcdQty"
            txtFlx.SelStart = 0
            txtFlx.SelLength = Len(txtFlx.Text)
            'txtFlx.SetFocus
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
    txtFlx.Tag = "skiplostfocus"
End Sub

Public Sub ComputeTotal()
    Dim i As Integer
    Dim TtlPrice As Currency, GrndTtl As Currency, DiscAmt As Currency, TtlDiscAmt As Currency
    Dim Pctd As Single, myPrice As Currency, TtlOrgPrice As Currency, GrndTtlOrgPrice As Currency
    
    With flxOrd
        GrndTtl = 0
        TtlDiscAmt = 0
        TtlOrgPrice = 0
        GrndTtlOrgPrice = 0
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 2)) > 0 And .TextMatrix(i, 7) <> "" Then
                TtlPrice = Val(.TextMatrix(i, 2)) * Val(Replace(.TextMatrix(i, 7), ",", ""))
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
            If Val(.TextMatrix(i, 13)) > 0 Then
                TtlOrgPrice = Val(.TextMatrix(i, 2)) * Val(Replace(.TextMatrix(i, 13), ",", ""))
                GrndTtlOrgPrice = GrndTtlOrgPrice + TtlOrgPrice
            End If
        Next i
        txtTtlDisc.Text = TtlDiscAmt
        txtPctDisc.Text = Val(Replace(GrndTtlOrgPrice, ",", "")) - Val(Replace(GrndTtl, ",", ""))
        txtTotal.Text = Format$(GrndTtl, "#,##0.00")
        txtGrandTotal.Text = Format$(GrndTtl - Val(Replace(txtLessDiscAmt.Text, ",", "")), "#,##0.00")
    End With
            
End Sub

Private Sub txtOrderno_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPSORDER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    If tabWS.Tag <> "" Then Exit Sub
    If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
    If KeyCode = vbKeyF9 And Not JustScroll Then Call cmdCashier_Click 'Key Accounts
End Sub

Private Sub txtPartNo_Change()
    Dim TtlPrice As Currency, DiscPrice As Currency
    Dim PctDisc As Single, myVat As Single, Qry As String
    
    On Error GoTo LocalErr
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
            flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
            flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
            flxOrd.TextMatrix(ActRow, 6) = ""   'Discount
            flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
            flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
            flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
            flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
            flxOrd.TextMatrix(ActRow, 14) = .TextMatrix(.Row, 9) 'COST
            flxOrd.TextMatrix(ActRow, 15) = .TextMatrix(.Row, 10) 'Remarks
            flxOrd.TextMatrix(ActRow, 16) = .TextMatrix(.Row, 11) 'Wholesale Price
            flxOrd.TextMatrix(ActRow, 17) = .TextMatrix(.Row, 12) 'Wholesale MAX discount
'            flxOrd.TextMatrix(ActRow, 18) = .TextMatrix(.Row, 14) 'Key Accounts Price
'            flxOrd.TextMatrix(ActRow, 19) = .TextMatrix(.Row, 15) 'Key Accounts MAX discount
            If Val(.TextMatrix(.Row, 9)) = 0 Then
                flxOrd.TextMatrix(ActRow, 7) = "" 'Price
                flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
            Else
                If ApplyWholeSalePriceInRetail And Val(.TextMatrix(.Row, 11)) > 0 Then
                    flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 11)), "###,##0.00") 'Dealer Price
                    flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 11)) 'Wholesale
                Else
                If ApplyKeyAcctsPriceInRetail And Val(.TextMatrix(.Row, 14)) > 0 Then
                    flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 14)), "###,##0.00") 'Key Accounts
                    flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 14)) 'Key Accounts
                Else
                    flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                    flxOrd.TextMatrix(ActRow, 13) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'SRP
                    
                    End If
                End If
            End If
            'Qty and SRP >0
            If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
                TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(Replace(flxOrd.TextMatrix(ActRow, 7), ",", ""))
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
            flxOrd.TextMatrix(ActRow, 15) = ""
            flxOrd.TextMatrix(ActRow, 16) = ""
            flxOrd.TextMatrix(ActRow, 17) = ""
            flxOrd.TextMatrix(ActRow, 18) = ""
            flxOrd.TextMatrix(ActRow, 19) = ""
            .Visible = False
            Call ComputeTotal
        End If
    End With
    Exit Sub
LocalErr:
    If Err.Number = 30009 Then
        flxParts.Visible = False
        txtPartNo.Visible = False
        If rsParts.State = adStateOpen Then rsParts.Close
        Timer1.Enabled = True
        MsgBox "Program is reloading Part Number" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
    Else
        flxParts.Visible = False
        txtPartNo.Visible = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Private Sub txtPartNo_GotFocus()
    If rsParts.State = adStateClosed Then Timer1.Enabled = True
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
            rsParts.MoveFirst
        Case vbKeyBack
            PrtSrch = True
            rsParts.MoveFirst
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    Dim eCol As Integer
      
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        
        'Call txtPartNo_LostFocus
        
        'If Val(flxOrd.TextMatrix(ActRow, 7)) = 0 And txtPartNo.Text <> "" Then
        '    MsgBox ("Sparepart Zero SRP" & vbCrLf & "Please contact your Manager"), vbCritical, "Zero SRP"
        '    For eCol = 1 To 14
        '        flxOrd.TextMatrix(ActRow, eCol) = ""
        '    Next eCol
        '    txtPartNo.Visible = False
        '    flxParts.Visible = False
        '    Exit Sub
        'End If
       
        flxOrd.SetFocus
        'If all fields are filled then addrow focus on next entry
        If Val(flxOrd.TextMatrix(ActRow, 2)) > Val(flxOrd.TextMatrix(ActRow, 1)) And flxOrd.TextMatrix(ActRow, 2) <> "" Then
            MsgBox "Qty Ordered must not be greater than Qty Available!", vbExclamation + vbOKOnly, "Validation"
            flxOrd.Row = ActRow
            flxOrd.Col = 2
            Call flxOrd_Click
            Exit Sub
        End If
        If flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
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
    Dim eCol As Integer, Qry As String
    Dim i As Long, j As Long
    Dim rsPers As New ADODB.Recordset
    
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    
    'Validate existing PartNo
    For i = 1 To flxOrd.Rows - 1
        If i <> ActRow Then
            If flxOrd.TextMatrix(i, 4) = txtPartNo.Text Then
                txtPartNo.Visible = False
                flxParts.Visible = False
                Ans = MsgBox("Part Number " & txtPartNo.Text & " already exist in row " & i, vbInformation + vbOKOnly, "Duplicate Part Number")
                For j = 1 To flxOrd.Cols - 1
                    flxOrd.TextMatrix(ActRow, j) = ""
                Next j
                flxOrd.Row = i
                flxOrd.Col = 4
                flxOrd.SetFocus
                txtPartNo.Tag = "PN exist"
                Exit Sub
            End If
        End If
    Next i
    
    If Val(flxOrd.TextMatrix(ActRow, 7)) = 0 And txtPartNo.Text <> "" Then
        MsgBox ("Sparepart Zero SRP" & vbCrLf & "Please contact your Manager"), vbCritical, "Zero SRP"
        For eCol = 1 To 14
           flxOrd.TextMatrix(ActRow, eCol) = ""
        Next eCol
        txtPartNo.Visible = False
        flxParts.Visible = False
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
    

'    Dim superseding As String
'    flxGridRef.TextMatrix(1, 0) = flxOrd.TextMatrix(ActRow, 4)
'    Qry = "Select sp.PARTNUMBER From mrcwarehouse.subspareparts as subsp Left Join mrcwarehouse.spareparts as sp ON sp.SPAREPARTID = subsp.SUBSPID Left Join mrcwarehouse.supplier as supp On supp.SUPPLIERCODE = sp.SUPPLIERCODE Left Join mrcwarehouse.partsmatrix as pm ON pm.PMATRIXID = sp.PMATRIXID Where subsp.SPAREPARTID =  ( SELECT mrcwarehouse.spareparts.SPAREPARTID FROM mrcwarehouse.spareparts WHERE  mrcwarehouse.spareparts.PARTNUMBER = '" & flxOrd.TextMatrix(ActRow, 4) & "');"
'    Qry = "sp_mrcretails.sp_supersedingSP('" & sysBranchCode & "','" & flxOrd.TextMatrix(ActRow, 4) & "')"
'
'     With rsPers
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            flxGridRef.Cols = 2
'            flxGridRef.Clear
'            flxGridRef.Cols = flxGridRef.Cols + 1
'            flxGridRef.TextMatrix(0, 0) = "PART NUMBER"
'            flxGridRef.TextMatrix(0, 2) = "QTY ON HAND"
'            Dim RowIndex, count As Integer
'            RowIndex = 1
'            count = 0
'            Do While Not .EOF
'                If Not IsNull(!PartNumber) Then
'                    flxGridRef.TextMatrix(RowIndex, 1) = !PartNumber
'                    flxGridRef.TextMatrix(RowIndex, 2) = !QtyOnHand
'                     count = count + 1
'                Else
'                    flxGridRef.TextMatrix(RowIndex, 1) = ""
'                    flxGridRef.TextMatrix(RowIndex, 2) = ""
'            End If
'                flxGridRef.TextMatrix(RowIndex, 0) = flxOrd.TextMatrix(ActRow, 4)
'                .MoveNext
'                RowIndex = RowIndex + 1
'
'            Loop
'              flxGridRef.TextMatrix(0, 1) = "SUPERSEDING PART NUMBER  (" & CStr(count) & ") RESULT"
'        Else
'            flxGridRef.TextMatrix(1, 1) = "NONE"
'            flxGridRef.TextMatrix(0, 1) = "SUPERSEDING PART NUMBER  (0)RESULT"
'            flxGridRef.TextMatrix(1, 2) = ""
'        End If
'        .Close
'       If count = 0 Then
'         flxGridRef.TextMatrix(1, 1) = "NONE"
'       End If
'
'    End With
'     flxGridRef.ColWidth(0) = flxGridRef.Width * 0.3
'    flxGridRef.ColWidth(1) = flxGridRef.Width * 0.4
'    flxGridRef.ColWidth(2) = flxGridRef.Width * 0.3
'    flxGridRef.ColAlignment(2) = flexAlignCenterCenter

End Sub

Private Sub txtOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtAddress.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        txtDate.SetFocus
    End If
End Sub

Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDate.SetFocus
End Sub

Private Sub txtPCCNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtBankCharge, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        cboPCPointsType.SetFocus
    End If
End Sub

Private Sub txtPCCNo_LostFocus()
    If txtPCCNo.Text <> "" Then
        If cboPCPointsType.Text <> "" Then
            Qry = "Select CPCID,CUSTOMERNAME,CUSTOMERPOINTS,CLUBPOINTS,CUSTOMERID,STATUS From custpromocardinfo Where customerno = '" & Trim(txtPCCNo.Text) & "'"
            With rsAnyUse
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    If !STATUS = "APPROVED" Then
                        txtPCCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
                        If cboPCPointsType.Text = "CUSTOMER POINTS" Then
                            txtPCCPoints.Text = IIf(IsNull(!CUSTOMERPOINTS), 0, !CUSTOMERPOINTS)
                        Else
                            txtPCCPoints.Text = IIf(IsNull(!CLUBPOINTS), 0, !CLUBPOINTS)
                        End If
                        myCardNo = Trim(txtPCCNo.Text)
                        myCardCustomerID = !CUSTOMERID
                    Else
                        MsgBox "Card Number Is Not Yet Approved!", vbInformation + vbOKOnly, "Validation"
                        txtPCCustomer.Text = ""
                        txtPCCPoints.Text = 0
                        myCardNo = ""
                        myCardCustomerID = 0
                        .Close
                        Exit Sub
                    End If
                Else
                    MsgBox "Card Number is not Available!", vbInformation + vbOKOnly, "Validation"
                    txtPCCustomer.Text = ""
                    txtPCCPoints.Text = 0
                    myCardNo = ""
                    myCardCustomerID = 0
                    .Close
                    Exit Sub
                End If
                .Close
            End With
        Else
            MsgBox "Unrecognized Points Type!, Please Select a Points Type", vbInformation + vbOKOnly
            cboPCPointsType.SetFocus
            Exit Sub
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
    '9/17/2015  NO DISPLAY OF COST USING X
    '04/05/16 IBALIK PERO C SIR JO LANG
    If Chr(KeyAscii) = "X" And UserID = 4 Then
        tmrQtyTip.Enabled = True
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxOrd.SetFocus
        'Call txtPrice_LostFocus
        If txtPrice.Tag = "XcdQty" Then txtPrice.Tag = "": Exit Sub
        'If all fields are filled then addrow focus on next entry
        If Val(txtPrice.Tag) > 0 And Val(txtPrice.Text) > 0 And Val(txtPrice.Tag) <> Val(txtPrice.Text) Then
            flxOrd.SetFocus
            flxOrd.Col = 7
            flxOrd.Row = ActRow
        ElseIf flxOrd.TextMatrix(ActRow, 5) <> "" And flxOrd.TextMatrix(ActRow, 2) <> "" Then
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

'Private Sub txtPrice_LostFocus()
'    Dim AmtDsc As Currency, LessAmt As Currency, Pctd As Double, myVat As Currency
'
'    If txtPrice.Tag = "skiplostfocus" Then
'        txtPrice.Tag = ""
'        Exit Sub
'    End If
'    With flxOrd
'        'Validate if cost=0 no price will be displayed
'        If Val(.TextMatrix(ActRow, 14)) = 0 Then
'            MsgBox "No available COST for this item." & vbCrLf & "Please inform your Manager.", vbOKOnly + vbExclamation, "Zero Cost"
'            flxOrd.TextMatrix(ActRow, 7) = ""
'            txtPrice.Visible = False
'            Call ComputeTotal
'            Exit Sub
'        Else
'            myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
'            AmtDsc = Val(.TextMatrix(ActRow, 14)) * myVat
'            If Val(Replace(txtPrice.Text, ",", "")) < Format(AmtDsc, "###0.00") Then
'                MsgBox "Invalid Price entered!" & vbCrLf & "Please follow the correct Price", vbOKOnly + vbExclamation, "Invalid Price"
'                flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
'                flxOrd.TextMatrix(ActRow, 6) = ""
'                txtPrice.Visible = False
'                Call ComputeTotal
'                Exit Sub
'            End If
'        End If
'
'        LessAmt = Val(.TextMatrix(ActRow, 13)) * (Val(.TextMatrix(ActRow, 12)) / 100)
'        AmtDsc = Val(.TextMatrix(ActRow, 13)) - LessAmt
'        If UserType <> "Administrator" Then
'            If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then
'                MsgBox "Price must not be less than maximum discounted Price: " & Format(AmtDsc, "#,##0.00") & "!", vbCritical + vbOKOnly, "Validation"
'                txtPrice.Tag = "XcdQty"
'                txtPrice.SelStart = 0
'                txtPrice.SelLength = Len(txtPrice.Text)
'                txtPrice.SetFocus
'                Exit Sub
'            End If
'        End If
'        If Val(Replace(txtPrice.Text, ",", "")) < Val(.TextMatrix(ActRow, 13)) Then
'            Pctd = (Val(.TextMatrix(ActRow, 13)) - Val(Replace(txtPrice.Text, ",", ""))) / Val(.TextMatrix(ActRow, 13))
'            If Val(Replace(txtPrice.Text, ",", "")) > 0 Then .TextMatrix(ActRow, 6) = Format(Pctd, "#0.000%")
'        End If
'        If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then 'Price is less than allowable discount
'            '9/17/2015  DO NOT ALLOW CHANGE PRICE BELOW DISCOUNT
'            flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
'            flxOrd.TextMatrix(ActRow, 6) = ""
'            .TextMatrix(ActRow, 3) = IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "", .TextMatrix(ActRow, 3) & "*")
'            Call ComputeTotal
'            txtPrice.Visible = False
'            MsgBox "Maximum discount is only " & Val(.TextMatrix(ActRow, 12)) & "%", vbInformation + vbOKOnly, "Invalid discount"
'            Exit Sub
'        Else
'            .TextMatrix(ActRow, 3) = Replace(.TextMatrix(ActRow, 3), "*", "")
'        End If
'    End With
'    flxOrd.TextMatrix(ActRow, 7) = Format(txtPrice.Text, "#,##0.00")
'    Call ComputeTotal
'    txtPrice.Visible = False
'
'End Sub

Private Sub txtPrice_LostFocus()
    Dim AmtDsc As Currency, LessAmt As Currency, Pctd As Double, myVat As Currency
    
    If txtPrice.Tag = "skiplostfocus" Then
        txtPrice.Tag = ""
        Exit Sub
    End If
    With flxOrd
        If Val(Replace(txtPrice.Text, ",", "")) >= Val(Replace(.TextMatrix(ActRow, 13), ",", "")) Then
            'Validate if cost=0 no price will be displayed
            If Val(.TextMatrix(ActRow, 14)) = 0 Then
                MsgBox "No available COST for this item." & vbCrLf & "Please inform your Manager.", vbOKOnly + vbExclamation, "Zero Cost"
                flxOrd.TextMatrix(ActRow, 7) = ""
                txtPrice.Visible = False
                Call ComputeTotal
                Exit Sub
            Else
                myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                AmtDsc = Val(.TextMatrix(ActRow, 14)) * myVat               '9/23/2015     allow less than Cost with VAT basta SALE
                If Val(Replace(txtPrice.Text, ",", "")) < Format(AmtDsc, "###0.00") And Mid(flxOrd.TextMatrix(ActRow, 5), 2, 4) <> "SALE" Then
                    MsgBox "Invalid Price entered!" & vbCrLf & "Please follow the correct Price", vbOKOnly + vbExclamation, "Invalid Price"
                    flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
                    flxOrd.TextMatrix(ActRow, 6) = ""
                    txtPrice.Visible = False
                    Call ComputeTotal
                    Exit Sub
                End If
            End If
            
            LessAmt = Val(.TextMatrix(ActRow, 13)) * (Val(.TextMatrix(ActRow, 12)) / 100) 'Get dscount = SRP x discount
            AmtDsc = Val(.TextMatrix(ActRow, 13)) - LessAmt 'Get discounted price  = SRP - discount
            If UserType <> "Administrator" Then
                If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then
                    MsgBox "Price must not be less than maximum discounted Price: " & Format(AmtDsc, "#,##0.00") & "!", vbCritical + vbOKOnly, "Validation"
                    txtPrice.Tag = "XcdQty"
                    txtPrice.SelStart = 0
                    txtPrice.SelLength = Len(txtPrice.Text)
                    txtPrice.SetFocus
                    Exit Sub
                End If
            End If              '9/23/2015      PRICE LESS THAN SRP AND DISCOUNT=0
            If Val(Replace(txtPrice.Text, ",", "")) < Val(.TextMatrix(ActRow, 13)) And Val(flxOrd.TextMatrix(ActRow, 12)) = 0 Then
    '            Pctd = (Val(.TextMatrix(ActRow, 13)) - Val(Replace(txtPrice.Text, ",", ""))) / Val(.TextMatrix(ActRow, 13))
    '            If Val(Replace(txtPrice.Text, ",", "")) > 0 Then .TextMatrix(ActRow, 6) = Format(Pctd, "#0.000%")
                flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
                flxOrd.TextMatrix(ActRow, 6) = ""
                .TextMatrix(ActRow, 3) = IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "", .TextMatrix(ActRow, 3) & "*")
                Call ComputeTotal
                txtPrice.Visible = False
                MsgBox "Discount is not available for this Part Number", vbInformation + vbOKOnly, "Invalid discount"
                Exit Sub
            End If
            
            If Not ((SysBranchID = 14 Or SysBranchID = 26) And chkClub.Value = 1 And flxOrd.TextMatrix(ActRow, 15) = "YAMAHA GENUINE") Then
              
                If UserID <> 4 Then '04/05/16 higher discount if user was CEO - advice sir jo
                
                    If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then 'Price is less than allowable discount
                        '9/17/2015  DO NOT ALLOW CHANGE PRICE BELOW DISCOUNT
                        flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
                        flxOrd.TextMatrix(ActRow, 6) = ""
                        .TextMatrix(ActRow, 3) = IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "", .TextMatrix(ActRow, 3) & "*")
                        Call ComputeTotal
                        txtPrice.Visible = False
                        MsgBox "Maximum discount is only " & Val(.TextMatrix(ActRow, 12)) & "%", vbInformation + vbOKOnly, "Invalid discount"
                        Exit Sub
                    Else
                        .TextMatrix(ActRow, 3) = Replace(.TextMatrix(ActRow, 3), "*", "")
                    End If
                End If
            
            Else ' for Yamaha Club Member
                LessAmt = Val(.TextMatrix(ActRow, 13)) * (Val(txtYClub.Text) / 100) 'Get dscount = SRP x discount
                AmtDsc = Val(.TextMatrix(ActRow, 13)) - LessAmt 'Get discounted price  = SRP - discount
                
                If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then 'Price is less than allowable discount
                    '9/17/2015  DO NOT ALLOW CHANGE PRICE BELOW DISCOUNT
                    flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
                    flxOrd.TextMatrix(ActRow, 6) = ""
                    .TextMatrix(ActRow, 3) = IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "", .TextMatrix(ActRow, 3) & "*")
                    Call ComputeTotal
                    txtPrice.Visible = False
                    MsgBox "Yamaha Club Member discount is only " & Val(txtYClub.Text) & "%", vbInformation + vbOKOnly, "Invalid discount"
                    Exit Sub
                Else
                    .TextMatrix(ActRow, 3) = Replace(.TextMatrix(ActRow, 3), "*", "")
                End If
            End If
            flxOrd.TextMatrix(ActRow, 7) = Format(txtPrice.Text, "#,##0.00")
            Call ComputeTotal
            txtPrice.Visible = False
        ElseIf Val(Replace(txtPrice.Text, ",", "")) < Val(Replace(.TextMatrix(ActRow, 13), ",", "")) Then
            If .TextMatrix(ActRow, 6) <> "" Then
                If Val(Replace(txtPrice.Text, ",", "")) < Val(Replace(.TextMatrix(ActRow, 7), ",", "")) Then
                    MsgBox "Price is less than allowable discount!.", vbInformation + vbOKOnly, "Warning"
                    txtPrice.Text = ""
                    txtPrice.Visible = False
                Else
                    flxOrd.TextMatrix(ActRow, 7) = Format(txtPrice.Text, "#,##0.00")
                    Call ComputeTotal
                    txtPrice.Visible = False
                End If
            Else
                MsgBox "New Price is less than the Original Price!.", vbInformation + vbOKOnly, "Warning"
                txtPrice.Text = ""
                txtPrice.Visible = False
                flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
            End If
            Exit Sub
        End If
    End With
End Sub


Sub UpdateNumOrder()
    Dim myOrder As New ADODB.Recordset, rsMyPvOrder As New ADODB.Recordset, myOrderDt As String
    Dim myPvOrder As Long, Qry As String, MyCrtOrder As Long, dts As Integer, myOrderNo As Integer
        'Query number of order today
        Qry = "SELECT " & _
            "spsalesorder.SPSORDERID " & _
        "From " & _
            "spsalesorder " & _
        "Where " & _
            "spsalesorder.ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "';"
            
            myOrder.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        If myOrder.RecordCount > 0 Then myOrder.Close: Exit Sub
            'Get previous max orderID and Add to the Setting
            dts = 1
            Do
                myOrderDt = Date - dts
                Qry = "SELECT " & _
                    "Max(spsalesorder.SPSORDERID) As LastSpID " & _
                "From " & _
                    "spsalesorder " & _
                "Where " & _
                    "spsalesorder.ORDERDATE = '" & Format(myOrderDt, "yyyy-mm-dd") & "';"
                 myPvOrder = Val(GetFieldValue(Qry))
                 
                 dts = dts + 1
            Loop Until myPvOrder <> 0
            
        'Save Last max orderID
        myOrder.Close
        LYRcon.Execute "Update settings SET `VALUE` = '" & myPvOrder & "', REMARKS = '" & Format(myOrderDt, "yyyy-mm-dd") & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "' " & _
            "Where settings.DESCRIPTION = 'SALES ORDER';"
End Sub

Private Sub TOT_PENDINGORDER(SP_ORDERDATE As String, SP_ID As Long)
    Dim rsPOrder As New ADODB.Recordset
    Dim QRY_ORDER As String, PENDINGORDER As Long, S_Order As Long, s_Col As Long
    
    QRY_ORDER = ""
    If AddNew Then
        QRY_ORDER = "SELECT SUM(spsorderdetail.QTY) AS TOTALQTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO " & _
            " From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
            " Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.INVOICENO IS NULL Group By " & _
            " spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO IS NULL;"
    Else
        QRY_ORDER = "SELECT SUM(spsorderdetail.QTY) AS TOTALQTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO " & _
            " From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
            " Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.ORDERNO <> " & Val(txtOrderno.Text) & " AND spsalesorder.INVOICENO IS NULL Group By " & _
            " spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO IS NULL;"
    End If
    
    PENDINGORDER = Val(GetFieldValue(QRY_ORDER))
    QRY_ORDER = "SELECT SUM(stocktransdtl.QTY) As TOTTRANS From stocktransfer RIGHT OUTER JOIN stocktransdtl ON (stocktransfer.TRANSFERID = stocktransdtl.TRANSFERID) " & _
                 " Where stocktransfer.`STATUS` = 'PENDING' AND stocktransdtl.SPAREPARTID = " & SP_ID & ";"
    TotPTrans = Val(GetFieldValue(QRY_ORDER))

    txtTotPen.Text = PENDINGORDER + TotPTrans
    If TotPTrans > 0 Then
        flxOrder.TextMatrix(0, 0) = "TRANS = " & TotPTrans
    End If
    With rsPOrder
    
        If AddNew Then
            QRY_ORDER = "SELECT spsorderdetail.QTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO, spsalesorder.FRONTLINER, spsalesorder.ORDERNO " & _
                "From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
                "Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.INVOICENO IS NULL;"
        Else
            QRY_ORDER = "SELECT spsorderdetail.QTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO, spsalesorder.FRONTLINER, spsalesorder.ORDERNO " & _
                "From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
                "Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.ORDERNO <> " & Val(txtOrderno.Text) & "  AND spsalesorder.INVOICENO IS NULL;"
        End If
       
       If .State = adStateOpen Then .Close
       .Open QRY_ORDER, LYRcon, adOpenKeyset, adLockOptimistic
       
       If .RecordCount > 0 Then
       
       flxOrder.Cols = .RecordCount + 1
       s_Col = 1
       .MoveFirst
           Do Until .EOF
               flxOrder.TextMatrix(0, s_Col) = "[ " & !ORDERNO & " ] " & !Frontliner & " = " & !Qty
               s_Col = s_Col + 1
               flxOrder.ColWidth(s_Col - 1) = flxOrder.Width * 0.25
               .MoveNext
           Loop
       End If
    End With
End Sub

Private Sub txtRPCardNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtBankCharge, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        cboRPPointsType.SetFocus
    End If
End Sub

Private Sub txtRPCardNo_LostFocus()
    If txtRPCardNo.Text <> "" Then
        If cboRPPointsType.Text <> "" Then
            Qry = "Select CPCID,CUSTOMERNAME,CUSTOMERPOINTS,CLUBPOINTS,CUSTOMERID From custpromocardinfo Where customerno = '" & Trim(txtRPCardNo.Text) & "' and STATUS = 'APPROVED'"
            With rsAnyUse
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    txtRPCustomer.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
                    If cboRPPointsType.Text = "CUSTOMER POINTS" Then
                        txtRPCardPoints.Text = IIf(IsNull(!CUSTOMERPOINTS), 0, !CUSTOMERPOINTS)
                    Else
                        txtRPCardPoints.Text = IIf(IsNull(!CLUBPOINTS), 0, !CLUBPOINTS)
                    End If
                    myCardNo = Trim(txtRPCardNo.Text)
                    myCardCustomerID = !CUSTOMERID
                Else
                    MsgBox "Card Number is Not Available!", vbInformation + vbOKOnly, "Validation"
                    txtRPCustomer.Text = ""
                    txtRPCardPoints.Text = 0
                    myCardNo = ""
                    myCardCustomerID = 0
                    .Close
                    Exit Sub
                End If
                .Close
            End With
        Else
            MsgBox "Unrecognized Points Type! Please Select a Points Type", vbInformation + vbOKOnly, "Validation"
            cboRPPointsType.SetFocus
            Exit Sub
        End If
    End If
End Sub
