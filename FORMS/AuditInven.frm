VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AuditInvntry 
   Caption         =   "Inventory Audit"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   11340
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxB4SubRpt 
      Height          =   735
      Left            =   3120
      TabIndex        =   58
      Top             =   9240
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.TextBox txtPrnQry 
      Height          =   285
      Left            =   5040
      TabIndex        =   57
      Top             =   8880
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtHomeID 
      Height          =   285
      Left            =   4080
      TabIndex        =   56
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSeqNo 
      Height          =   285
      Left            =   3120
      TabIndex        =   55
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab tabLoc 
      Height          =   3255
      Left            =   4080
      TabIndex        =   50
      Top             =   10080
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "B1002032"
      TabPicture(0)   =   "AuditInven.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdLocCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdLocOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "flxLoc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSFlexGridLib.MSFlexGrid flxLoc 
         Height          =   2175
         Left            =   240
         TabIndex        =   52
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin lvButton.lvButtons_H cmdLocOK 
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   2805
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
         Image           =   "AuditInven.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdLocCancel 
         Height          =   375
         Left            =   1800
         TabIndex        =   54
         Top             =   2805
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
         Image           =   "AuditInven.frx":05B6
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "Please select a Spare Part:"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   8640
      TabIndex        =   49
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtBatchID 
      Height          =   285
      Left            =   7800
      TabIndex        =   48
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFind 
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
      Left            =   8880
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   9600
      TabIndex        =   30
      Top             =   10320
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   2040
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   7800
      TabIndex        =   29
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9000
      TabIndex        =   28
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboParts 
      Height          =   315
      Left            =   11040
      TabIndex        =   27
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   480
      TabIndex        =   26
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
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   360
      TabIndex        =   24
      Tag             =   "XYF"
      Top             =   7440
      Width           =   10695
      Begin VB.TextBox txtTotal 
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
         Height          =   285
         Left            =   8880
         TabIndex        =   47
         Tag             =   "XYF"
         Top             =   810
         Width           =   1695
      End
      Begin VB.TextBox txtFinal 
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
         Height          =   285
         Left            =   8880
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   525
         Width           =   1695
      End
      Begin VB.TextBox txtTemp 
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
         Height          =   285
         Left            =   8880
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks 
         Height          =   765
         Left            =   1080
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL RECORDS"
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
         Left            =   6720
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   847
         Width           =   1980
      End
      Begin VB.Label lblFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "FINAL RECORDS"
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
         Left            =   6720
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   562
         Width           =   1980
      End
      Begin VB.Label lblTemp 
         Alignment       =   1  'Right Justify
         Caption         =   "TEMPORARY RECORDS"
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
         Left            =   6720
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   277
         Width           =   1980
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
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   360
      TabIndex        =   23
      Text            =   "Qty. Purchased: 0"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   10200
      TabIndex        =   22
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid flxInv 
      Height          =   5055
      Left            =   360
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   2400
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "^    |<PART NUMBER      |<DESCRIPTION          |>PRICE         |<LOCATION     |^QTY.         |^ TIME         |^DATE    "
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
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1245
      Left            =   360
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   720
      Width           =   10695
      Begin VB.TextBox txtDateFin 
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
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtAuditor 
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
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtStatus 
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
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
      Begin VB.TextBox txtSubReport 
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
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   520
         Width           =   1815
      End
      Begin VB.TextBox txtDateStart 
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
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtBatchNo 
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE FINISH:"
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
         Left            =   6840
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   870
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "AUDITOR'S NAME:"
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
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   870
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE START:"
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
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   550
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "BATCH NO."
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   230
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "SUB REPORTS:"
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
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   550
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "REPORT STATUS"
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
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   230
         Width           =   1815
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   15
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
      Image           =   "AuditInven.frx":0B50
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   16
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Discrepancy"
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
      Image           =   "AuditInven.frx":10EA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   17
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Mgr's Report"
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
      Image           =   "AuditInven.frx":1684
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7500
      TabIndex        =   18
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
      Image           =   "AuditInven.frx":1C1E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   19
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
      Image           =   "AuditInven.frx":21B8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   10500
      TabIndex        =   20
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
      Image           =   "AuditInven.frx":2752
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   21
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Edit Report"
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
      Image           =   "AuditInven.frx":2CEC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   31
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "S&ummary"
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
      Image           =   "AuditInven.frx":3286
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   9
      Left            =   11520
      TabIndex        =   32
      Top             =   1320
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "AuditInven.frx":3820
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   8
      Left            =   11520
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Sub Report"
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
      Image           =   "AuditInven.frx":3DBA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   10
      Left            =   11520
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Main Report"
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
      Image           =   "AuditInven.frx":4354
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   7740
      TabIndex        =   40
      Tag             =   "XYF"
      Top             =   2010
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      Caption         =   "&Find"
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
      Image           =   "AuditInven.frx":48EE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdInsertSP 
      Height          =   375
      Left            =   6600
      TabIndex        =   41
      Tag             =   "XYF"
      Top             =   2010
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      Caption         =   "&Insert SP"
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
      Image           =   "AuditInven.frx":4E88
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11580
      m_MDI_Height    =   9360
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTORY AUDIT"
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
      Left            =   9180
      TabIndex        =   13
      Tag             =   "TRF"
      Top             =   450
      Width           =   1755
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
      TabIndex        =   11
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
      TabIndex        =   12
      Tag             =   "TLF"
      Top             =   475
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTORY AUDIT"
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
      Left            =   9210
      TabIndex        =   14
      Tag             =   "TRF"
      Top             =   480
      Width           =   1755
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
   End
End
Attribute VB_Name = "AuditInvntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim ActRow As Integer, ActCol As Integer, fRow As Integer, fCol As Integer, LocRow As Integer
Dim AddNew As Boolean, LockFlx As Boolean, DispOnly As Boolean, PrtSrch As Boolean, PrtBln As Boolean, CustSrch As Boolean

Private Sub ShowButtons(MyType As Integer)
    Dim i As Integer
    
    If MyType = 1 Then
        For i = 0 To 5
            cmdTrans(i).Left = cmdTrans(0).Width * i
            cmdTrans(i).Visible = True
        Next i
        For i = 6 To 10
            cmdTrans(i).Visible = False
        Next i
    ElseIf MyType = 2 Then 'Save Cancel
        For i = 0 To 10
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(6).Left = 0
        cmdTrans(7).Left = cmdTrans(6).Width
        cmdTrans(6).Visible = True
        cmdTrans(7).Visible = True
    ElseIf MyType = 3 Then 'Sub Report, Print Preview, Main Report, Summary, Exit
        For i = 0 To 10
            cmdTrans(i).Visible = False
        Next i
        For i = 0 To 2
            cmdTrans(8 + i).Top = 0
            cmdTrans(8 + i).Left = cmdTrans(8).Width * i
            cmdTrans(8 + i).Visible = True
        Next i
        cmdTrans(4).Left = cmdTrans(8).Width * 3
        cmdTrans(4).Visible = True
        cmdTrans(5).Left = cmdTrans(8).Width * 4
        cmdTrans(5).Visible = True
    ElseIf MyType = 4 Then 'Sub Report, Print Preview, Main Report, Summary, Exit
        For i = 0 To 10
            cmdTrans(i).Visible = False
        Next i
        For i = 0 To 1
            cmdTrans(9 + i).Top = 0
            cmdTrans(9 + i).Left = cmdTrans(9).Width * i
            cmdTrans(9 + i).Visible = True
        Next i
        cmdTrans(4).Left = cmdTrans(9).Width * 2
        cmdTrans(4).Visible = True
        cmdTrans(5).Left = cmdTrans(9).Width * 3
        cmdTrans(5).Visible = True
    End If
End Sub

Private Sub ClearEntries()
    txtStatus.Text = ""
    txtSubReport.Text = ""
    txtAuditor.Text = ""
    txtBatchNo.Text = ""
    txtDateStart.Text = ""
    txtDateFin.Text = ""
    txtFind.Text = ""
    txtRemarks.Text = ""
    txtTemp.Text = ""
    txtFinal.Text = ""
    txtTotal.Text = ""
    txtPrnCopy.Text = ""
    txtBatchID.Text = ""
    flxInv.Rows = 1
    flxInv.Rows = 2
    txtPrnQry.Text = ""
    txtHomeID.Text = ""
    txtSeqNo.Text = ""
    flxB4SubRpt.Rows = 1
    flxB4SubRpt.Rows = 2
End Sub

Private Sub LockData(bln As Boolean)
    'Frame1.Enabled = Not Bln
    Frame2.Enabled = Not bln
    cmdInsertSP.Enabled = Not bln
    LockFlx = bln
End Sub

Private Sub ShowMgrReport(MyID As Integer)
    Dim rsAInv As New ADODB.Recordset
    Dim rsAIdtl As New ADODB.Recordset
    Dim Qry As String, Tmp As Integer
    
    'Qry = "SELECT AUDITINVEN.* From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & MyId & "));"
    Qry = "SELECT AUDITINVEN.* From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & MyID & ";"
    
    With rsAInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtStatus.Text = !STATUS
            txtSubReport.Text = IIf(IsNull(!BatchMember), 0, !BatchMember)
            txtAuditor.Text = !Auditor
            txtBatchNo.Text = !BatchNo
            txtDateStart.Text = !DateStart
            txtDateFin.Text = IIf(IsNull(!DateEnd), "", Format(!DateEnd, "mm/dd/yyyy"))
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtSeqNo.Text = !SEQNO
            txtHomeID.Text = !HOMEID
            txtBatchID.Text = MyID
        End If
        .Close
    End With
    With flxInv
        .Rows = 1
        .Rows = 2
        .TextMatrix(0, 5) = "QTY [PHY]"
        .TextMatrix(0, 6) = "QTY [REC]"
        .TextMatrix(0, 7) = "VARIANCE"
    End With
    'txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST]*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.DISCREPANCY) <> 0) And ((AUDITINVDTLS.BATCHID) = " & txtBatchID.Text & ")) ORDER BY AUDITINVDTLS.ROWNO;"
    'mySQL
    txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.DISCREPANCY <> 0 And AUDITINVDTLS.BATCHID = " & txtBatchID.Text & " ORDER BY AUDITINVDTLS.ROWNO;"
    txtPrnQry.Tag = "MGR"
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST]*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.DISCREPANCY) <> 0) And ((AUDITINVDTLS.BATCHID) = " & txtBatchID.Text & ")) ORDER BY AUDITINVDTLS.ROWNO;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.DISCREPANCY <> 0 And AUDITINVDTLS.BATCHID = " & txtBatchID.Text & " ORDER BY AUDITINVDTLS.ROWNO;"
    
    With rsAIdtl
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxInv.Rows = 1
            Do While Not .EOF
                Tmp = Tmp + IIf(IsNull(!LASTUPDATE), 1, 0)
                flxInv.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!COSTWITHVAT, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & Format(!PhyQty, "#,##0") & vbTab & IIf(IsNull(!QTYONPC), "", !QTYONPC) & vbTab & IIf(IsNull(!DISCREPANCY), "", !DISCREPANCY)
                .MoveNext
            Loop
            txtTemp.Text = Format(Tmp, "#,##0") & Space(7)
            txtFinal.Text = Format(.RecordCount - Tmp, "#,##0") & Space(7)
            txtTotal.Text = Format(.RecordCount, "#,##0") & Space(7)
        Else
            txtTemp.Text = "0" & Space(7)
            txtFinal.Text = "0" & Space(7)
            txtTotal.Text = "0" & Space(7)
        End If
        .Close
    End With
End Sub

Public Sub ShowRecord(MyID As Integer)
    Dim rsAInv As New ADODB.Recordset
    Dim rsAIdtl As New ADODB.Recordset
    Dim Qry As String, Tmp As Integer
    
    'Qry = "SELECT AUDITINVEN.* From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT AUDITINVEN.* From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & MyID & ";"
    
    With rsAInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtStatus.Text = !STATUS
            txtSubReport.Text = IIf(IsNull(!BatchMember), 0, !BatchMember)
            txtAuditor.Text = !Auditor
            txtBatchNo.Text = !BatchNo
            txtDateStart.Text = !DateStart
            txtDateFin.Text = IIf(IsNull(!DateEnd), "", Format(!DateEnd, "mm/dd/yyyy"))
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtSeqNo.Text = !SEQNO
            txtHomeID.Text = !HOMEID
            txtBatchID.Text = MyID
        End If
        .Close
    End With
    With flxInv
        .Rows = 1
        .Rows = 2
        .TextMatrix(0, 5) = "QTY"
        .TextMatrix(0, 6) = "TIME"
        .TextMatrix(0, 7) = "DATE"
    End With
    'txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST]*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.BATCHID) = " & MyId & ")) ORDER BY AUDITINVDTLS.ROWNO;"
    txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.BATCHID = " & MyID & " ORDER BY AUDITINVDTLS.ROWNO;"

    txtPrnQry.Tag = ""
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST]*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.BATCHID) = " & MyId & ")) ORDER BY AUDITINVDTLS.ROWNO;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
          "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.BATCHID = " & MyID & " ORDER BY AUDITINVDTLS.ROWNO;"
    
    With rsAIdtl
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxInv.Rows = 1
            Do While Not .EOF
                Tmp = Tmp + IIf(IsNull(!LASTUPDATE), 1, 0)
                flxInv.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!COSTWITHVAT, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & Format(!PhyQty, "#,###") & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "HH:MM ampm")) & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "mm/dd/yyyy")) & vbTab & IIf(IsNull(!LASTUPDATE), "", !LASTUPDATE) & vbTab & IIf(IsNull(!QTYONPC), "", !QTYONPC) & vbTab & IIf(IsNull(!DISCREPANCY), "", !DISCREPANCY) & vbTab & !SparepartID
                .MoveNext
            Loop
            txtTemp.Text = Format(Tmp, "#,##0") & Space(7)
            txtFinal.Text = Format(.RecordCount - Tmp, "#,##0") & Space(7)
            txtTotal.Text = Format(.RecordCount, "#,##0") & Space(7)
        Else
            txtTemp.Text = "0" & Space(7)
            txtFinal.Text = "0" & Space(7)
            txtTotal.Text = "0" & Space(7)
        End If
        .Close
    End With
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Private Sub LoadSpareParts()
    Dim Qry As String
        
    'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST]*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = flxInv.ColWidth(4)
        .ColWidth(0) = 0
        .ColWidth(1) = 0    'SPid
        .ColWidth(2) = flxInv.ColWidth(4)   'PartNo
        .ColWidth(3) = 0    'Description
        .ColWidth(4) = 0    'Cost with VAT
        .ColWidth(5) = 0    'Location
    End With
End Sub
Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxInv
        For i = fRow To .Rows - 1
            For j = fCol To 7
                If InStr(UCase$(.TextMatrix(i, j)), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 7:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 7, i + 1, i)
                    fCol = IIf(j = 7, 1, j + 1)
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

Private Sub cmdInsertSP_Click()
    SPfind.MySource = "AUDITINVEN"
    SPfind.Show vbModal
End Sub

Private Sub cmdLocCancel_Click()
    cmdTrans(6).Enabled = True
    cmdTrans(7).Enabled = True
    flxInv.Enabled = True
    Frame2.Enabled = True
    txtFind.Locked = False
    cmdFind.Enabled = True
    cmdInsertSP.Enabled = True
    tabLoc.Top = 15000
    tabLoc.Visible = False
    flxInv.SetFocus
    flxInv.Row = ActRow
    flxInv.Col = 5
End Sub

Private Sub cmdLocOK_Click()
    Dim i As Integer
        
    With flxLoc
        'Validate if Partnumber already exist
        For i = 1 To flxInv.Rows - 1
            If flxInv.TextMatrix(i, 11) = .TextMatrix(LocRow, 1) Then
                MsgBox .TextMatrix(.Row, 2) & " already exist at Row " & i, vbExclamation + vbOKOnly, "Location"
                Exit Sub
            End If
        Next i
        flxInv.TextMatrix(ActRow, 11) = .TextMatrix(LocRow, 1)
        flxInv.TextMatrix(ActRow, 1) = .TextMatrix(LocRow, 2)
        flxInv.TextMatrix(ActRow, 2) = .TextMatrix(LocRow, 3)
        flxInv.TextMatrix(ActRow, 3) = .TextMatrix(LocRow, 4)
        flxInv.TextMatrix(ActRow, 4) = .TextMatrix(LocRow, 5)
    End With
    Call cmdLocCancel_Click
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim rsAud As New ADODB.Recordset
    Dim Ans As Integer, i As Integer, SubRptID As Integer
    Dim myMsg As String, PrnVal(0 To 12) As String, Audi As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=94));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=94;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            
            Call ClearEntries
            Call LockData(False)
            Call ShowButtons(2)
            AddNew = True
            txtBatchNo.Text = GetBatchNo
            txtDateStart.Text = Date
            txtStatus.Text = "ON PROCESS"
            txtSubReport.Text = "0"
            Audi = InputBox("Please type Name of Auditor:", "In-Charge")
            If Audi = "" Then
                MsgBox "No Auditor Specified... Transaction aborted!", vbExclamation + vbOKOnly, "Cancelled"
                Call ClearEntries
                Call LockData(True)
                Call ShowButtons(1)
                Exit Sub
            End If
            txtAuditor.Text = UCase$(Audi)
            
            'save records to table
            'rsAud.Open "AuditInven", LYRcon, adOpenKeyset, adLockOptimistic
            'rsAud.AddNew
            'rsAud!BatchNo = txtBatchNo.Text
            'rsAud!Auditor = txtAuditor.Text
            'rsAud!DateStart = Date
            'rsAud!Status = txtStatus.Text
            'rsAud!SeqNo = 1
            'ACCESS
            'rsAud!LockEntry = True
            'mySQL
            'rsAud!LockEntry = 1
            'rsAud!UserID = UserID
            'rsAud.Update
            'txtBatchID.Text = rsAud!BATCHID
            'rsAud.Close
            
            LYRcon.Execute "INSERT INTO auditinven(BATCHNO, AUDITOR, DATESTART, `STATUS`, " & _
                "SEQNO, LOCKENTRY, USERID)VALUES('" & txtBatchNo.Text & "', '" & txtAuditor.Text & "', " & _
                "'" & Format(Date, "yyyy-mm-dd") & "', '" & txtStatus.Text & "', 1, 1, '" & UserID & "');"
            txtBatchID.Text = GetFieldValue("Select auditinven.BATCHID FROM auditinven WHERE auditinven.BatchNo = '" & txtBatchNo.Text & "';")
        
        Case 1  'Edit
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=94));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=94;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            'Validation for Editing
            If txtStatus.Text = "FINAL" Then
                MsgBox "Editing is not allowed because the Report is already Final.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            End If
            If Val(txtBatchID.Text) = 0 Then
                MsgBox "No Report available for editing.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            End If
            'Identify if Report is Being used by Other User
            'Qry = "SELECT AUDITINVEN.LOCKENTRY From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            'mySQL
            Qry = "SELECT AUDITINVEN.LOCKENTRY From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            myMsg = GetFieldValue(Qry)
            If myMsg = "True" Then  '
                'Qry = "SELECT AUDITINVEN.USERID From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
                'mySQL
                Qry = "SELECT AUDITINVEN.USERID From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
                i = Val(GetFieldValue(Qry))
                'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS USERNAME, EMPLOYEES.EMPID " & _
                      "From EMPLOYEES WHERE (((EMPLOYEES.EMPID)=" & i & "));"
                'mySQL
                Qry = "EMPLOYEES.EMPID, concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS USERNAME, EMPLOYEES.EMPID " & _
                      "From EMPLOYEES WHERE EMPLOYEES.EMPID=" & i & ";"
                Audi = GetFieldValue(Qry)
                MsgBox "Unavailable to Edit Report. Currently used by " & Audi, vbInformation + vbOKOnly, "Edit Deneid"
                Exit Sub
            End If
            'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = True, AUDITINVEN.USERID = " & UserID & " WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            'mySQL
            LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = 1, AUDITINVEN.USERID = " & UserID & " WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"

            AddNew = False
            Call LockData(False)
            Call ShowButtons(2)
            
        Case 2  'Display Discrepancy
            'validate
            If Val(txtBatchID.Text) = 0 Then
                MsgBox "No available report to process!", vbExclamation + vbOKOnly, "Inventory Audit"
                Exit Sub
            End If
            If txtStatus.Text = "FINAL" And Val(txtSubReport.Text) > 0 Then
                'Qry = "SELECT AUDITINVEN.BATCHID From AUDITINVEN " & _
                      "WHERE (((AUDITINVEN.HOMEID)=" & txtHomeID.Text & ") AND ((AUDITINVEN.SEQNO)=" & Val(txtSeqNo.Text) + 1 & "));"
                'mySQL
                Qry = "SELECT AUDITINVEN.BATCHID From AUDITINVEN " & _
                      "WHERE AUDITINVEN.HOMEID=" & txtHomeID.Text & " AND AUDITINVEN.SEQNO=" & Val(txtSeqNo.Text) + 1 & ";"
                
                SubRptID = Val(GetFieldValue(Qry))
                Call ShowRecord(SubRptID)
                Call ShowButtons(1)
                Exit Sub
            End If
            With flxInv
                .Rows = 1
                .Rows = 2
                .TextMatrix(0, 5) = "QTY"
                .TextMatrix(0, 6) = "TIME"
                .TextMatrix(0, 7) = "DATE"
            End With
            'txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST]*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
                  "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.DISCREPANCY) <> 0) And ((AUDITINVDTLS.BATCHID) = " & txtBatchID.Text & ")) ORDER BY AUDITINVDTLS.ROWNO;"
            'mySQL
            txtPrnQry.Text = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE " & _
                  "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.DISCREPANCY <> 0 And AUDITINVDTLS.BATCHID = " & txtBatchID.Text & " ORDER BY AUDITINVDTLS.ROWNO;"
            txtPrnQry.Tag = ""
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST]*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
                  "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((AUDITINVDTLS.DISCREPANCY) <> 0) And ((AUDITINVDTLS.BATCHID) = " & txtBatchID.Text & ")) ORDER BY AUDITINVDTLS.ROWNO;"
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION, AUDITINVDTLS.PHYQTY, AUDITINVDTLS.LASTUPDATE, AUDITINVDTLS.QTYONPC, AUDITINVDTLS.DISCREPANCY, AUDITINVDTLS.SPAREPARTID " & _
                  "FROM AUDITINVDTLS INNER JOIN SPAREPARTS ON AUDITINVDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where AUDITINVDTLS.DISCREPANCY <> 0 And AUDITINVDTLS.BATCHID = " & txtBatchID.Text & " ORDER BY AUDITINVDTLS.ROWNO;"
            
            With rsAud
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    flxInv.Rows = 1
                    .MoveFirst
                    Do While Not .EOF
                        flxInv.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!COSTWITHVAT, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & Format(!PhyQty, "#,##0") & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "HH:MM ampm")) & vbTab & IIf(IsNull(!LASTUPDATE), "", Format(!LASTUPDATE, "mm/dd/yyyy")) & vbTab & IIf(IsNull(!LASTUPDATE), "", !LASTUPDATE) & vbTab & IIf(IsNull(!QTYONPC), "", !QTYONPC) & vbTab & IIf(IsNull(!DISCREPANCY), "", !DISCREPANCY) & vbTab & !SparepartID
                        .MoveNext
                    Loop
                    txtTemp.Text = "0" & Space(7)
                    txtFinal.Text = Format(.RecordCount, "#,###") & Space(7)
                    txtTotal.Text = Format(.RecordCount, "#,###") & Space(7)
                Else
                    flxInv.Rows = 1
                    flxInv.Rows = 2
                    txtTemp.Text = ""
                    txtFinal.Text = ""
                    txtTotal.Text = ""
                End If
                .Close
            End With
            Call ShowButtons(3)
        Case 3  'Managers Report
            'validate
            If Val(txtBatchID.Text) = 0 Then
                MsgBox "No available report to process!", vbExclamation + vbOKOnly, "Inventory Audit"
                Exit Sub
            End If
            'Prompt for Password
            Password.MySource = "AUDITINV"
            Password.Show vbModal
            If cmdTrans(3).Tag <> "" Then
                cmdTrans(3).Tag = ""
                Exit Sub
            End If
            'Qry = "SELECT AUDITINVEN.BATCHID From AuditInven Where (((AuditInven.HOMEID) = " & txtHomeID.Text & ")) ORDER BY AUDITINVEN.SEQNO DESC;"
            'mySQL
            Qry = "SELECT AUDITINVEN.BATCHID From AuditInven Where AuditInven.HOMEID = " & txtHomeID.Text & " ORDER BY AUDITINVEN.SEQNO DESC;"
            'SubRptID = Val(GetFieldValue(Qry))
            Call ShowMgrReport(SubRptID)
            Call ShowButtons(4)
            
        Case 4  'Summary
            AuditInvSum.Show
            Unload Me
        
        Case 5  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
            Exit Sub
            
        Case 6 'Save
            myMsg = ValidateInvAdj
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Adjustment?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveAuditInv
            MsgBox "Inventory report Successfully Saved!", vbInformation + vbOKOnly, "Save"
            
            Call LockData(True)
            Call ShowButtons(1)
        Case 7  'Cancel
            Ans = MsgBox("Do you want to Cancel transaction?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                'remove record
                'LYRcon.Execute "DELETE AUDITINVEN.* From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & Val(txtBatchID.Text) & "));"
                'mySQL
                LYRcon.Execute "DELETE From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & Val(txtBatchID.Text) & ";"
                'LYRcon.Execute "DELETE From AUDITINVDTLS WHERE (((AUDITINVDTLS.BATCHID)=" & Val(txtBatchID.Text) & "));"
                'mySQL
                LYRcon.Execute "DELETE From AUDITINVDTLS WHERE AUDITINVDTLS.BATCHID=" & Val(txtBatchID.Text) & ";"
                Call ClearEntries
            Else
                Qry = "SELECT From AUDITINVDTLS WHERE (((AUDITINVDTLS.BATCHID)=" & Val(txtBatchID.Text) & "));"
                If Not CtrlNumExist(Qry) And flxB4SubRpt.TextMatrix(1, 6) = "REVERT" Then 'flxInv.Rows = 2 And flxInv.TextMatrix(1, 1) = "" And
                    'LYRcon.Execute "DELETE From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & Val(txtBatchID.Text) & "));"
                    'mySQL
                    LYRcon.Execute "DELETE From AUDITINVDTLS WHERE AUDITINVDTLS.BATCHID=" & Val(txtBatchID.Text) & ";"
                    'LYRcon.Execute "DELETE AUDITINVDTLS.* From AUDITINVDTLS WHERE (((AUDITINVDTLS.BATCHID)=" & Val(txtBatchID.Text) & "));"
                    'mySQL
                    LYRcon.Execute "DELETE From AUDITINVDTLS WHERE AUDITINVDTLS.BATCHID=" & Val(txtBatchID.Text) & ";"
                    
                    'Revert Previous Report
                    With flxB4SubRpt
                        'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = [AUDITINVEN].[BATCHMEMBER]-1 WHERE (((AUDITINVEN.BATCHID)=" & .TextMatrix(1, 4) & "));"
                        'mySQL
                        LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = AUDITINVEN.BATCHMEMBER-1 WHERE AUDITINVEN.BATCHID=" & .TextMatrix(1, 4) & ";"
                        
                        'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.STATUS = '" & .TextMatrix(1, 3) & "', AUDITINVEN.DATEEND = " & IIf(IsDate(.TextMatrix(1, 2)), "#" & .TextMatrix(1, 2) & "#", "Null") & ", AUDITINVEN.DISCREPANCY = 0, AUDITINVEN.SEQNO = " & Val(.TextMatrix(1, 5)) & _
                                       " WHERE (((AUDITINVEN.BATCHID)=" & Val(.TextMatrix(1, 0)) & "));"
                        'mySQL
                        LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.STATUS = '" & .TextMatrix(1, 3) & "', AUDITINVEN.DATEEND = " & IIf(IsDate(.TextMatrix(1, 2)), "'" & Format(.TextMatrix(1, 2), "yyyy-mm-dd") & "'", "Null") & ", AUDITINVEN.DISCREPANCY = 0, AUDITINVEN.SEQNO = " & Val(.TextMatrix(1, 5)) & _
                                       " WHERE AUDITINVEN.BATCHID=" & Val(.TextMatrix(1, 0)) & ";"
                        flxB4SubRpt.Rows = 1
                        flxB4SubRpt.Rows = 2
                    End With
                End If
                'Old Comment LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = False WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & ")"
                Call ClearEntries
            End If
            Call ShowButtons(1)
            Call LockData(True)
        Case 8  'Sub Report
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=94));"
            'mySQL
            
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=94;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            'Check if there are existing sub-reports and display if none create new
            If flxInv.Rows = 2 And flxInv.TextMatrix(1, 11) = "" Then
                MsgBox "No Discrepancies available!", vbInformation + vbOKOnly, "Sub-Report"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Create Sub-Report?", vbQuestion + vbYesNo, "Sub-Report")
            If Ans = vbNo Then Exit Sub
            'Store Current Value of Active Report
            With flxB4SubRpt
                .TextMatrix(1, 0) = txtBatchID.Text
                .TextMatrix(1, 1) = txtDateStart.Text
                .TextMatrix(1, 2) = txtDateFin.Text
                .TextMatrix(1, 3) = txtStatus.Text
                .TextMatrix(1, 4) = txtHomeID.Text
                .TextMatrix(1, 5) = txtSeqNo.Text
                .TextMatrix(1, 6) = "REVERT"
                'BatchMember = Batchmember - 1
                'Descripancy = 0
            End With
            'Update Home's Batch Members
            'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = [AUDITINVEN].[BATCHMEMBER]+1 WHERE (((AUDITINVEN.BATCHID)=" & txtHomeID.Text & "));"
            'mySQL
            LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = AUDITINVEN.BATCHMEMBER+1 WHERE AUDITINVEN.BATCHID=" & txtHomeID.Text & ";"
            'If Val(txtBatchID.Text) <> Val(txtHomeID.Text) Then LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = [AUDITINVEN].[BATCHMEMBER]+1 WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            If Val(txtBatchID.Text) <> Val(txtHomeID.Text) Then LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.BATCHMEMBER = AUDITINVEN.BATCHMEMBER+1 WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            
            
            'Finalize Previous Batch
            'Qry = "SELECT AUDITINVEN.STATUS From AUDITINVEN WHERE (((AUDITINVEN.STATUS)='FINAL') AND ((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            'mySQL
            Qry = "SELECT AUDITINVEN.STATUS From AUDITINVEN WHERE AUDITINVEN.STATUS='FINAL' AND AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            If CtrlNumExist(Qry) Then   'if Prev Batch already final update only Discrepancy else include Status and Date Finish
                'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.DISCREPANCY = " & flxInv.Rows - 1 & " WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
                'mySQL
                LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.DISCREPANCY = " & flxInv.Rows - 1 & " WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            Else
                'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.STATUS = 'FINAL', AUDITINVEN.DATEEND = #" & Date & "#, AUDITINVEN.DISCREPANCY = " & flxInv.Rows - 1 & _
                " WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
                'mySQL
                LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.STATUS = 'FINAL', AUDITINVEN.DATEEND = '" & Format(Date, "yyyy-mm-dd") & "', AUDITINVEN.DISCREPANCY = " & flxInv.Rows - 1 & _
                    " WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
            
            End If
            'Qry = "SELECT AUDITINVEN.BATCHNO From AUDITINVEN WHERE (((AUDITINVEN.BATCHID)=" & txtHomeID.Text & "));"
            'mySQL
            Qry = "SELECT AUDITINVEN.BATCHNO From AUDITINVEN WHERE AUDITINVEN.BATCHID=" & txtHomeID.Text & ";"
            myMsg = GetFieldValue(Qry)
            txtBatchNo.Text = myMsg & "-" & ConvertToRoman(txtSeqNo.Text)
            txtDateStart.Text = Date
            txtStatus.Text = "ON PROCESS"
            txtSubReport.Text = "0"
            txtSeqNo.Text = Val(txtSeqNo.Text) + 1
            
            'Create New Batch
            'rsAud.Open "AuditInven", LYRcon, adOpenKeyset, adLockOptimistic
            'rsAud.AddNew
            'rsAud!BATCHNO = txtBatchNo.Text
            'rsAud!AUDITOR = txtAuditor.Text
            'rsAud!DATESTART = Date
            'rsAud!Status = txtStatus.Text
            'rsAud!SEQNO = Val(txtSeqNo.Text)
            'rsAud!HOMEID = txtHomeID.Text
            'Access
            'rsAud!LockEntry = True
            'mySQL
            'rsAud!LockEntry = 1
            'rsAud!UserID = UserID
            'rsAud.Update
            'txtBatchID.Text = rsAud!BATCHID
            'rsAud.Close
            
            LYRcon.Execute "INSERT INTO auditinven(BATCHNO, AUDITOR, DATESTART, `STATUS`, SEQNO, HOMEID, LOCKENTRY, " & _
                "USERID)VALUES('" & txtBatchNo.Text & "', '" & txtAuditor.Text & "', '" & Format(Date, "yyyy-mm-dd") & "', " & _
                "'" & txtStatus.Text & "', " & Val(txtSeqNo.Text) & ", " & txtHomeID.Text & ", 1, " & UserID & ");"
            txtBatchID.Text = GetFieldValue("Select auditinven.BATCHID FROM auditinven WHERE auditinven.BatchNo = '" & txtBatchNo.Text & "';")
      
            
            Call LockData(False)
            Call ShowButtons(2)
            AddNew = False
        Case 9  'Print Preview
            If txtPrnQry.Tag = "" Then
                PRNmc.QryPrn = txtPrnQry.Text
                PRNmc.ParamValue1 = txtBatchNo.Text
                PRNmc.ParamValue2 = txtAuditor.Text   'cboType.Text
                PRNmc.ParamValue3 = IIf(txtStatus.Text = "FINAL", "", "Inventory Audit is still ON-PROCESS") & " " & txtRemarks.Text
                PRNmc.MySource = "AUDITINVEN"
            Else
                PRNmc.QryPrn = txtPrnQry.Text
                PRNmc.ParamValue1 = txtBatchNo.Text
                PRNmc.ParamValue2 = txtAuditor.Text   'cboType.Text
                '***Access
                'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.REMARKS)='MANAGER') AND ((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE'));"
                'mySQL
                Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.REMARKS='MANAGER' AND SETTINGS.DESCRIPTION='EMP DEFAULT VALUE';"
                Audi = GetFieldValue(Qry)
                PRNmc.ParamValue3 = Audi
                PRNmc.ParamValue4 = IIf(txtStatus.Text = "FINAL", "", "Inventory Audit is still ON-PROCESS") & " " & txtRemarks.Text
                Ans = MsgBox("Do you want to print report with Cost?", vbQuestion + vbYesNo, "Sub-Report")
                If Ans = vbYes Then
                    PRNmc.MySource = "AUDINVMGRCOST"
                Else
                    PRNmc.MySource = "AUDINVMGR"
                End If
            End If
            PRNmc.Show
            
        Case 10 'Main Report
            'Display Main Report
            Call ShowButtons(1)
            Call ShowRecord(Val(txtHomeID.Text))
            
    End Select
End Sub

Private Sub SaveAuditInv()
    Dim i As Integer, Tmp As Integer

    With flxInv
        If AddNew Then
            Tmp = 0
            For i = 1 To .Rows - 1
                If Val(flxInv.TextMatrix(i, 11)) > 0 Then
                    If flxInv.TextMatrix(i, 8) = "" Then    'Record Not yet updated
                        '***Access
                        'LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, ROWNO ) " & _
                            "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", " & i & ");"
                        
                        'mySQL
                        LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, ROWNO ) " & _
                            "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", " & i & ");"
                        Tmp = Tmp + 1
                    End If
                End If
            Next i
            '***access
            'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.HOMEID = " & txtBatchID.Text & " WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
            'mySQL
            LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.HOMEID = " & txtBatchID.Text & " WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
        Else
            'LYRcon.Execute "DELETE AUDITINVDTLS.* From AUDITINVDTLS WHERE (((AUDITINVDTLS.BATCHID)=" & txtBatchID.Text & "));"
            'mySQL
            LYRcon.Execute "DELETE From AUDITINVDTLS WHERE AUDITINVDTLS.BATCHID=" & txtBatchID.Text & ";"
            For i = 1 To .Rows - 1
                If Val(flxInv.TextMatrix(i, 11)) > 0 Then
                    If flxInv.TextMatrix(i, 8) = "" Then    'Record Not yet updated
                        LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, QTYONPC, DISCREPANCY, ROWNO ) " & _
                                "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", " & Val(flxInv.TextMatrix(i, 9)) & ", " & Val(flxInv.TextMatrix(i, 10)) & ", " & i & ");"
                          Tmp = Tmp + 1
                    Else
                        'LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, LASTUPDATE, QTYONPC, DISCREPANCY, ROWNO ) " & _
                                "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", #" & flxInv.TextMatrix(i, 8) & "#, " & Val(flxInv.TextMatrix(i, 9)) & ", " & Val(flxInv.TextMatrix(i, 10)) & ", " & i & ");"
                        'mySQL
                        LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, LASTUPDATE, QTYONPC, DISCREPANCY, ROWNO ) " & _
                                "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(i, 11) & ", " & Val(Replace(flxInv.TextMatrix(i, 5), ",", "")) & ", '" & Format(flxInv.TextMatrix(i, 8), "yyyy-mm-dd") & "', " & Val(flxInv.TextMatrix(i, 9)) & ", " & Val(flxInv.TextMatrix(i, 10)) & ", " & i & ");"
                    End If
                End If
            Next i
        End If
        txtTemp.Text = Format(Tmp, "#,##0") & Space(7)
        txtFinal.Text = Format((.Rows - 1) - Tmp, "#,##0") & Space(7)
        txtTotal.Text = Format((.Rows - 1), "#,##0") & Space(7)
    End With
    'LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = False WHERE (((AUDITINVEN.BATCHID)=" & txtBatchID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE AUDITINVEN SET AUDITINVEN.LOCKENTRY = 0 WHERE AUDITINVEN.BATCHID=" & txtBatchID.Text & ";"
    
End Sub

Private Sub flxInv_Click()
    Dim Ans As Integer, MyQty As Integer
    Dim Qry As String
    ActRow = flxInv.Row
    ActCol = flxInv.Col
    
    If ActCol = 0 Or ActRow = 0 Or LockFlx Then Exit Sub
    Select Case ActCol
        Case 1 'PartNumber
            With txtPartNo
                If flxInv.TextMatrix(ActRow, 1) <> "" Then flxInv.TextMatrix(ActRow, 12) = flxInv.TextMatrix(ActRow, 1) & "|" & flxInv.TextMatrix(ActRow, 2) & "|" & flxInv.TextMatrix(ActRow, 3) & "|" & flxInv.TextMatrix(ActRow, 4) & "|" & flxInv.TextMatrix(ActRow, 5) & "|" & flxInv.TextMatrix(ActRow, 6) & "|" & flxInv.TextMatrix(ActRow, 7) & "|" & flxInv.TextMatrix(ActRow, 8) & "|" & flxInv.TextMatrix(ActRow, 9) & "|" & flxInv.TextMatrix(ActRow, 10) & "|" & flxInv.TextMatrix(ActRow, 11)
                .Top = flxInv.CellTop + flxInv.Top
                .Left = flxInv.CellLeft + flxInv.Left
                .Width = flxInv.CellWidth
                .Text = flxInv.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 4  'Location
            With txtLocation
                If flxInv.TextMatrix(ActRow, 4) <> "" Then .Tag = flxInv.TextMatrix(ActRow, 4)
                .Top = flxInv.CellTop + flxInv.Top
                .Left = flxInv.CellLeft + flxInv.Left
                .Width = flxInv.CellWidth
                .Text = flxInv.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 5    'qty
            With txtQty
                If flxInv.TextMatrix(ActRow, 5) <> "" Then .Tag = flxInv.TextMatrix(ActRow, 5) & "|" & flxInv.TextMatrix(ActRow, 6) & "|" & flxInv.TextMatrix(ActRow, 7) & "|" & flxInv.TextMatrix(ActRow, 8) & "|" & flxInv.TextMatrix(ActRow, 9) & "|" & flxInv.TextMatrix(ActRow, 10)
                .Top = flxInv.CellTop + flxInv.Top
                .Left = flxInv.CellLeft + flxInv.Left
                .Width = flxInv.CellWidth
                .Text = flxInv.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 6  'Time
            If flxInv.TextMatrix(ActRow, 5) = "" Or Val(flxInv.TextMatrix(ActRow, 11)) = 0 Then Exit Sub
            Ans = MsgBox("Do you want to Finalize update on Physical Quantity?", vbQuestion + vbYesNo, "Final")
            If Ans = vbNo Then Exit Sub
            'Qry = "SELECT SPAREPARTS.QTYONHAND From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & flxInv.TextMatrix(ActRow, 11) & "));"
            'mySQL
            Qry = "SELECT SPAREPARTS.QTYONHAND From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID=" & flxInv.TextMatrix(ActRow, 11) & ";"
          
            MyQty = Val(GetFieldValue(Qry))
            flxInv.TextMatrix(ActRow, 9) = MyQty
            flxInv.TextMatrix(ActRow, 10) = Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) - MyQty
            flxInv.TextMatrix(ActRow, 8) = Now
            flxInv.TextMatrix(ActRow, 7) = Format(Now, "mm/dd/yyyy")
            flxInv.TextMatrix(ActRow, 6) = Format(Now, "hh:mm ampm")
            'Qry = "SELECT AUDITINVDTLS.BATCHID From AUDITINVDTLS WHERE (((AUDITINVDTLS.BATCHID)=" & txtBatchID.Text & ") AND ((AUDITINVDTLS.SPAREPARTID)=" & flxInv.TextMatrix(ActRow, 11) & "));"
            'mySQL
            Qry = "SELECT AUDITINVDTLS.BATCHID From AUDITINVDTLS WHERE AUDITINVDTLS.BATCHID=" & txtBatchID.Text & " AND AUDITINVDTLS.SPAREPARTID=" & flxInv.TextMatrix(ActRow, 11) & ";"
            
            If CtrlNumExist(Qry) Then
                'LYRcon.Execute "UPDATE AUDITINVDTLS SET AUDITINVDTLS.PHYQTY = " & Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) & ", AUDITINVDTLS.LASTUPDATE = #" & flxInv.TextMatrix(ActRow, 8) & "#, AUDITINVDTLS.QTYONPC = " & flxInv.TextMatrix(ActRow, 9) & ", AUDITINVDTLS.DISCREPANCY = " & flxInv.TextMatrix(ActRow, 10) & ", AUDITINVDTLS.ROWNO = " & ActRow & _
                " WHERE (((AUDITINVDTLS.BATCHID)=" & txtBatchID.Text & ") AND ((AUDITINVDTLS.SPAREPARTID)=" & flxInv.TextMatrix(ActRow, 11) & "));"
            'mySQL
            LYRcon.Execute "UPDATE AUDITINVDTLS SET AUDITINVDTLS.PHYQTY = " & Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) & ", AUDITINVDTLS.LASTUPDATE = '" & Format(flxInv.TextMatrix(ActRow, 8), "yyyy-mm-dd") & "', AUDITINVDTLS.QTYONPC = " & flxInv.TextMatrix(ActRow, 9) & ", AUDITINVDTLS.DISCREPANCY = " & flxInv.TextMatrix(ActRow, 10) & ", AUDITINVDTLS.ROWNO = " & ActRow & _
                " WHERE AUDITINVDTLS.BATCHID=" & txtBatchID.Text & " AND AUDITINVDTLS.SPAREPARTID=" & flxInv.TextMatrix(ActRow, 11) & ";"
            Else
                'LYRcon.Execute "INSERT INTO AUDITINVDTLS ( BATCHID, SPAREPARTID, PHYQTY, LASTUPDATE, QTYONPC, DISCREPANCY, ROWNO ) " & _
                              "VALUES (" & txtBatchID.Text & ", " & flxInv.TextMatrix(ActRow, 11) & ", " & Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) & ", #" & flxInv.TextMatrix(ActRow, 8) & "#, " & Val(flxInv.TextMatrix(ActRow, 9)) & ", " & Val(flxInv.TextMatrix(ActRow, 10)) & ", " & ActRow & ");"
            'mySQL
                LYRcon.Execute "INSERT INTO AUDITINVDTLS(BATCHID, SPAREPARTID, PHYQTY, LASTUPDATE, QTYONPC, DISCREPANCY, ROWNO) " & _
                              "VALUES(" & txtBatchID.Text & ", " & flxInv.TextMatrix(ActRow, 11) & ", " & Val(Replace(flxInv.TextMatrix(ActRow, 5), ",", "")) & ", '" & Format(flxInv.TextMatrix(ActRow, 8), "yyyy-mm-dd") & "', " & Val(flxInv.TextMatrix(ActRow, 9)) & ", " & Val(flxInv.TextMatrix(ActRow, 10)) & ", " & ActRow & ");"
            End If
            MsgBox "Record successfully updated!", vbOKOnly + vbInformation, flxInv.TextMatrix(ActRow, 1)
            flxInv.SetFocus
            flxInv.Col = 6
            If ActRow < flxInv.Rows - 1 Then
                flxInv.Row = ActRow + 1
            Else
                flxInv.Row = ActRow
            End If
    End Select
End Sub

Private Sub flxInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxInv_Click
End Sub

Private Sub flxInv_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If LockFlx Then Exit Sub
    If Button = vbRightButton And ActCol = 0 Then
        Call PopupMenu(mnuADJ)
    End If
End Sub

Private Sub flxInv_Scroll()
    txtPartNo.Visible = False
    txtLocation.Visible = False
    txtQty.Visible = False
End Sub

Private Sub flxLoc_DblClick()
    Call cmdLocOK_Click
End Sub

Private Sub flxLoc_EnterCell()
    LocRow = flxLoc.Row
End Sub

Private Sub flxLoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdLocOK_Click
End Sub

Private Sub flxParts_Click()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxInv.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3) 'Description
        flxInv.TextMatrix(ActRow, 3) = Format$(Val(.TextMatrix(.Row, 4)), "#,##0.00") 'Price
        flxInv.TextMatrix(ActRow, 4) = .TextMatrix(.Row, 5) 'Location
        flxInv.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 1) 'SPid
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
        flxInv.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3) 'Description
        flxInv.TextMatrix(ActRow, 3) = Format$(Val(.TextMatrix(.Row, 4)), "#,##0.00") 'Price
        flxInv.TextMatrix(ActRow, 4) = .TextMatrix(.Row, 5) 'Location
        flxInv.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 1) 'SPid
    End With
End Sub

Private Sub flxParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        Tmp = Split(flxInv.TextMatrix(ActRow, 12), "|")
        For i = LBound(Tmp) To UBound(Tmp)
            flxInv.TextMatrix(ActRow, i + 1) = Tmp(i)
        Next i
        txtPartNo.Visible = False
        flxParts.Visible = False
        flxInv.SetFocus
        flxInv.Col = 1
    End If

End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        flxInv.SetFocus
        flxInv.Row = ActRow
        flxInv.Col = 5
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Call LockData(True)
    Call ShowButtons(1)
    cmdA.Top = 15000
    Timer1.Enabled = True
    MotoRep.sBar.Panels(6).Text = "Spare Parts Inventory Auditing"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub

Private Sub mnuAdd_Click()
    If flxInv.TextMatrix(ActRow, 11) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxInv.SetFocus
        flxInv.Col = ActCol
        Exit Sub
    Else
        If ActRow = flxInv.Rows - 1 Then flxInv.AddItem ""
        flxInv.SetFocus
        flxInv.Col = 1
        flxInv.Row = flxInv.Rows - 1
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If ActRow = 1 And flxInv.Rows = 2 Then
            flxInv.Rows = 1
            flxInv.Rows = 2
        Else
            flxInv.RemoveItem ActRow
        End If
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxInv
        .ColWidth(0) = 350
        .ColWidth(1) = (.Width - 380) * 0.18
        .ColWidth(2) = (.Width - 380) * 0.21
        .ColWidth(3) = (.Width - 380) * 0.11
        .ColWidth(4) = (.Width - 380) * 0.15
        .ColWidth(5) = (.Width - 380) * 0.11
        .ColWidth(6) = (.Width - 380) * 0.11
        .ColWidth(7) = (.Width - 380) * 0.11
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
    End With
    With flxLoc
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = .Width - 200
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColAlignment(2) = 4
    End With
End Sub

Function GetBatchNo() As String
    Dim rsAdjCntr As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    'Qry = "SELECT Max(Val([BATCHNO])) AS MyBatch From AUDITINVEN WHERE (((AUDITINVEN.SEQNO)=1));"
    'NO Availabhle Funtion of Val from mySQL
    'mySQL
    Qry = "SELECT Max(round(AUDITINVEN.BATCHNO)) AS MyBatch From AUDITINVEN WHERE AUDITINVEN.SEQNO=1;"

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
    GetBatchNo = MyNo
End Function

Function ValidateInvAdj() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""

    If flxInv.TextMatrix(flxInv.Rows - 1, 11) = "" Then
        If flxInv.Rows <= 2 Then
            vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Adjusted Spare Parts"
        Else
            flxInv.RemoveItem flxInv.Rows - 1
        End If
    End If
    With flxInv
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 11)) = 0 Then
                vMsg = vMsg & vbCrLf & "Invalid Part Number at row " & i
            End If
        Next i
    End With
    ValidateInvAdj = vMsg
End Function

Private Sub Timer1_Timer()
    Call LoadSpareParts
    Timer1.Enabled = False
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxInv.SetFocus
        flxInv.Col = 5
        flxInv.Row = ActRow
    End If
End Sub

Private Sub txtLocation_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtLocation.Text = Empty
        flxInv.TextMatrix(ActRow, 4) = txtLocation.Tag
        txtLocation.Tag = "skiplostfocus"
        flxInv.SetFocus
        flxInv.Col = 4
        txtLocation.Visible = False
    End If
End Sub
Private Sub txtLocation_LostFocus()
    Dim rsLoc As New ADODB.Recordset
    Dim Qry As String
    
    If txtLocation.Tag = "skiplostfocus" Then
        txtLocation.Tag = ""
        Exit Sub
    End If
    If txtLocation.Text = "" Then
        MsgBox "No Location specified!", vbExclamation + vbOKOnly, "Location"
        txtLocation.Visible = False
        Exit Sub
    End If
    'Find Location in Database
    'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format([COST]*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS Where (((SPAREPARTS.Location) Like '%" & txtLocation.Text & "%')) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, Format(SPAREPARTS.COST*1.12,'Fixed') AS COSTWITHVAT, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS Where SPAREPARTS.Location Like '%" & txtLocation.Text & "%' ORDER BY SPAREPARTS.PARTNUMBER;"

    With rsLoc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then
            MsgBox "No Spare Part assigned on this location!", vbExclamation + vbOKOnly, "Location"
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            flxInv.TextMatrix(ActRow, 1) = !PartNumber
            flxInv.TextMatrix(ActRow, 2) = !Description
            flxInv.TextMatrix(ActRow, 3) = IIf(IsNull(!COSTWITHVAT), "", Format(!COSTWITHVAT, "#,##0.00"))
            flxInv.TextMatrix(ActRow, 4) = txtLocation.Text
            flxInv.TextMatrix(ActRow, 11) = !SparepartID
        ElseIf .RecordCount > 1 Then
            .MoveFirst
            cmdTrans(6).Enabled = False
            cmdTrans(7).Enabled = False
            flxInv.Enabled = False
            Frame2.Enabled = False
            txtFind.Locked = True
            cmdFind.Enabled = False
            cmdInsertSP.Enabled = False
            flxLoc.Rows = 0
            tabLoc.Top = flxInv.Top + 600
            tabLoc.Left = (Me.Width - tabLoc.Width) / 2
            tabLoc.Caption = txtLocation.Text
            tabLoc.Visible = True
            tabLoc.ZOrder
            Do While Not .EOF
                flxLoc.AddItem vbTab & !SparepartID & vbTab & !PartNumber & vbTab & !Description & vbTab & IIf(IsNull(!COSTWITHVAT), "", Format(!COSTWITHVAT, "#,##0.00")) & vbTab & txtLocation.Text
                .MoveNext
            Loop
            LocRow = 1
            flxLoc.SetFocus
        End If
        .Close
    End With
    txtLocation.Visible = False
End Sub

Private Sub txtPartNo_Change()
    Dim TtlPrice As Currency, DiscPrice As Currency
    Dim PctDisc As Single, i As Integer
    
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
            flxInv.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3) 'Description
            flxInv.TextMatrix(ActRow, 3) = Format$(Val(.TextMatrix(.Row, 4)), "#,##0.00") 'Price
            flxInv.TextMatrix(ActRow, 4) = .TextMatrix(.Row, 5) 'Location
            flxInv.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 1) 'SPid
            For i = 5 To 10
                flxInv.TextMatrix(ActRow, i) = ""   'Clear Previous Entry
            Next i
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            flxInv.TextMatrix(ActRow, 2) = ""
            flxInv.TextMatrix(ActRow, 3) = ""
            flxInv.TextMatrix(ActRow, 4) = ""
            flxInv.TextMatrix(ActRow, 11) = ""
            .Visible = False
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
            flxParts.Tag = "skiplostfocus"
            flxParts.Visible = True
            flxParts.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            If flxInv.TextMatrix(ActRow, 12) <> "" Then
                Tmp = Split(flxInv.TextMatrix(ActRow, 12), "|")
                For i = LBound(Tmp) To UBound(Tmp)
                    flxInv.TextMatrix(ActRow, i + 1) = Tmp(i)
                Next i
            Else
                For i = 1 To 12
                    flxInv.TextMatrix(ActRow, i) = ""
                Next i
            End If
            flxParts.Tag = "skiplostfocus"
            flxInv.SetFocus
            flxInv.Col = 1
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
        flxInv.SetFocus
        flxInv.Row = ActRow
        flxInv.Col = 5
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    If flxParts.Tag = "skiplostfocus" Then
        flxParts.Tag = ""
        Exit Sub
    End If
    If flxInv.TextMatrix(ActRow, 11) = "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxParts.Visible = False
        Exit Sub
    End If
    flxInv.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxParts.Visible = False
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    If KeyCode = vbKeyEscape Then
        txtQty.Text = Empty
        If txtQty.Tag <> "" Then
            Tmp = Split(txtQty.Tag, "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxInv.TextMatrix(ActRow, i + 5) = Tmp(i)
            Next i
        Else
            For i = 5 To 10
                flxInv.TextMatrix(ActRow, i) = ""
            Next i
        End If
        txtQty.Tag = "skiplostfocus"
        flxInv.SetFocus
        flxInv.Col = 5
        txtQty.Visible = False
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxInv
            .SetFocus
            If .TextMatrix(ActRow, 11) <> "" Then
                If ActRow = .Rows - 1 Then .AddItem ""
                .Row = .Rows - 1
                .Col = 1
            Else
                .Row = ActRow
                .Col = 5
            End If
        End With
    End If
End Sub

Private Sub txtQty_LostFocus()
    Dim i As Integer
    
    If txtQty.Tag = "skiplostfocus" Then
        txtQty.Tag = ""
        Exit Sub
    End If
    If txtQty.Text <> txtQty.Tag Then
        For i = 5 To 10
            flxInv.TextMatrix(ActRow, i) = ""
        Next i
    End If
    flxInv.TextMatrix(ActRow, 5) = txtQty.Text
    txtQty.Visible = False
End Sub

Private Function ConvertToRoman(pstrDecimalNumber As String) As String
'--------------------------------------------------------------------------------------

    Const strPOS_VAL    As String = "IXCM"
    Const strFIVE_VAL   As String = "VLD"
    
    Dim strRoman        As String
    Dim strCurrRomanPos As String
    Dim strLetter1      As String
    Dim strLetter2      As String
    Dim intCurrPos      As Integer
    Dim intDigit        As Integer
    Dim intDigitPos     As Integer
    
    intCurrPos = 1
    strRoman = ""
    
    For intDigitPos = Len(pstrDecimalNumber) To 1 Step -1
        intDigit = Val(Mid$(pstrDecimalNumber, intDigitPos, 1))
        strCurrRomanPos = Mid$(strPOS_VAL, intCurrPos, 1)
        Select Case intDigit
            Case 9
                strLetter1 = strCurrRomanPos
                strLetter2 = Mid$(strPOS_VAL, intCurrPos + 1, 1)
            Case Is > 4
                strLetter1 = Mid$(strFIVE_VAL, intCurrPos, 1)
                strLetter2 = String$(intDigit - 5, strCurrRomanPos)
            Case 4
                strLetter1 = strCurrRomanPos
                strLetter2 = Mid$(strFIVE_VAL, intCurrPos, 1)
            Case Else
                strLetter1 = String$(intDigit, strCurrRomanPos)
                strLetter2 = ""
        End Select
        strRoman = strLetter1 & strLetter2 & strRoman
        intCurrPos = intCurrPos + 1
    Next

    ConvertToRoman = strRoman

End Function

