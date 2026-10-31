VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SPInsXst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Existing Spare Parts"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16845
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SPInsXst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   16845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Branch Database Update"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   65
      Top             =   8760
      Width           =   16575
      Begin VB.Label lblBrnCon2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity Formula"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   67
         Top             =   480
         Width           =   15900
      End
      Begin VB.Label lblBrnCon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity Formula"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   16140
      End
   End
   Begin TabDlg.SSTab tabProcessing 
      Height          =   975
      Left            =   6000
      TabIndex        =   70
      Tag             =   "xyf"
      Top             =   3600
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1720
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Processing"
      TabPicture(0)   =   "SPInsXst.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProcess"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Analyzing Inventory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   72
         Tag             =   "xyf"
         Top             =   600
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   240
         Picture         =   "SPInsXst.frx":05A6
         Tag             =   "xyf"
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "PLEASE WAIT...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   71
         Tag             =   "xyf"
         Top             =   360
         Width           =   3255
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   5775
      Left            =   45
      TabIndex        =   68
      Top             =   2400
      Width           =   16635
      _ExtentX        =   29342
      _ExtentY        =   10186
      _Version        =   393216
      RowHeightMin    =   300
      HighLight       =   0
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   8280
      Top             =   135
   End
   Begin MSComCtl2.FlatScrollBar fScrol 
      Height          =   195
      Left            =   5400
      TabIndex        =   53
      ToolTipText     =   "Click to Select Branch"
      Top             =   840
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   2
      Arrows          =   65536
      Min             =   1
      Max             =   16
      Orientation     =   1245185
      Value           =   1
   End
   Begin VB.Frame fmAnotation 
      Caption         =   "EXPLANATORY NOTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   120
      TabIndex        =   40
      Top             =   8760
      Visible         =   0   'False
      Width           =   16560
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOQ ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   11520
         TabIndex        =   73
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity Formula"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   12240
         TabIndex        =   64
         Top             =   270
         Width           =   2940
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOQ = Branch Total SQ - (Branch QH+Warehouse QH)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   11520
         TabIndex        =   63
         Top             =   585
         Width           =   4470
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "(6) months SQ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9960
         TabIndex        =   62
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6mSQ = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   9240
         TabIndex        =   61
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Cost W/Out Vat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9960
         TabIndex        =   60
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CWOV = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   9240
         TabIndex        =   59
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Qty."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   58
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOQ = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4800
         TabIndex        =   56
         Top             =   585
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Stock Transfer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   55
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ST = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4800
         TabIndex        =   54
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   645
         TabIndex        =   52
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label6 
         Caption         =   " Supplier BACK-ORDER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   600
         TabIndex        =   51
         Top             =   600
         Width           =   1920
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sales Quantity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3015
         TabIndex        =   50
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3015
         TabIndex        =   49
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7740
         TabIndex        =   48
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Quantity Taking"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7740
         TabIndex        =   47
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   150
         TabIndex        =   46
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BO ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   585
         Width           =   405
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SQ = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2520
         TabIndex        =   44
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2520
         TabIndex        =   43
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7305
         TabIndex        =   42
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QT ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7305
         TabIndex        =   41
         Top             =   585
         Width           =   405
      End
   End
   Begin MSComCtl2.UpDown UpDRef 
      Height          =   315
      Left            =   1800
      TabIndex        =   36
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Max             =   1
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   8280
      Visible         =   0   'False
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtOrder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   19
      Top             =   9600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdReplenish 
      Height          =   495
      Left            =   11520
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "Stocks for Replenish"
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
      Image           =   "SPInsXst.frx":182D0
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtSuppCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   17
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   375
      Left            =   12600
      TabIndex        =   7
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Insert"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPInsXst.frx":1886A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptHide 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Hide Option"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ImgAlign        =   3
      Image           =   "SPInsXst.frx":18E04
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptShow 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Show Option"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ImgAlign        =   3
      Image           =   "SPInsXst.frx":1939E
      cBack           =   -2147483633
   End
   Begin VB.Frame FrameOpt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   45
      TabIndex        =   3
      Top             =   720
      Width           =   16635
      Begin VB.ComboBox cboClass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   75
         Tag             =   "XYF"
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox chkResize 
         Height          =   345
         Left            =   11520
         Picture         =   "SPInsXst.frx":19938
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   405
      End
      Begin VB.ComboBox cbomos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "SPInsXst.frx":1A202
         Left            =   8880
         List            =   "SPInsXst.frx":1A204
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   720
         Width           =   735
      End
      Begin MSComCtl2.UpDown OptrSQ 
         Height          =   285
         Left            =   15360
         TabIndex        =   35
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown OptrBO 
         Height          =   285
         Left            =   13080
         TabIndex        =   34
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtQH 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13335
         TabIndex        =   31
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.UpDown updBO 
         Height          =   285
         Left            =   13800
         TabIndex        =   27
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBO"
         BuddyDispid     =   196647
         OrigLeft        =   9840
         OrigTop         =   600
         OrigRight       =   10095
         OrigBottom      =   855
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtBO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13320
         TabIndex        =   26
         Text            =   "0"
         Top             =   615
         Width           =   480
      End
      Begin MSComCtl2.UpDown updSQ 
         Height          =   285
         Left            =   16080
         TabIndex        =   25
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSQ"
         BuddyDispid     =   196648
         OrigLeft        =   11280
         OrigTop         =   255
         OrigRight       =   11535
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSQ 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15600
         TabIndex        =   24
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.DTPicker SDfrom 
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   39498
      End
      Begin VB.CheckBox chkAllSupp 
         Caption         =   "All Suppliers:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   420
         Width           =   1335
      End
      Begin VB.ComboBox cboSupplier 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   360
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker SDto 
         Height          =   315
         Left            =   6240
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   39498
      End
      Begin MSComCtl2.UpDown OptrQH 
         Height          =   285
         Left            =   13080
         TabIndex        =   30
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updQH 
         Height          =   285
         Left            =   13800
         TabIndex        =   32
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtQH"
         BuddyDispid     =   196646
         OrigLeft        =   9015
         OrigTop         =   255
         OrigRight       =   9270
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Order Mos.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7920
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   780
         Width           =   885
      End
      Begin VB.Label lblOprQH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   12600
         TabIndex        =   33
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblOprSQ 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   14880
         TabIndex        =   29
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblOprBO 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   12600
         TabIndex        =   28
         Top             =   615
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "SQ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   14400
         TabIndex        =   23
         ToolTipText     =   "SALES QTY."
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "BO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   12120
         TabIndex        =   22
         ToolTipText     =   "BACK ORDER"
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "QH:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   12120
         TabIndex        =   21
         ToolTipText     =   "QTY-ON-HAND"
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Date To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Date From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Classification:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   420
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   135
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "SPInsXst.frx":1A206
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   13980
      TabIndex        =   8
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPInsXst.frx":1A7A0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   15360
      TabIndex        =   9
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPInsXst.frx":1AD3A
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPH 
      Height          =   510
      Left            =   45
      TabIndex        =   39
      Top             =   1845
      Width           =   16635
      _ExtentX        =   29342
      _ExtentY        =   900
      _Version        =   393216
      Cols            =   82
      FixedRows       =   2
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdEvaluate 
      Height          =   495
      Left            =   6480
      TabIndex        =   69
      Tag             =   "xyf"
      Top             =   120
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   873
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
      Image           =   "SPInsXst.frx":1B2D4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdReplenish2 
      Height          =   495
      Left            =   14160
      TabIndex        =   74
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "Stocks for Replenish"
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
      Image           =   "SPInsXst.frx":1B86E
      cBack           =   -2147483633
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
      Left            =   90
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblRefSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIND PARTS DESC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   315
      Width           =   1485
   End
   Begin VB.Menu mnuSP 
      Caption         =   "SP"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAct 
         Caption         =   "set to ""ACTIVE"""
      End
      Begin VB.Menu mnuInAct 
         Caption         =   "set to ""IN-ACTIVE"""
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMark 
         Caption         =   "Mark"
      End
      Begin VB.Menu mnuUnmark 
         Caption         =   "Unmark"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkAll 
         Caption         =   "Mark All"
      End
      Begin VB.Menu mnuUnmarkAll 
         Caption         =   "UnMark All"
      End
   End
   Begin VB.Menu mnurg 
      Caption         =   "RG"
      Visible         =   0   'False
      Begin VB.Menu mnuresize 
         Caption         =   "Resize Grid"
      End
   End
End
Attribute VB_Name = "SPInsXst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSupplier As New ADODB.Recordset
Dim ShowOpt As Boolean, SuppSrch As Boolean
Dim ActRow As Integer, ActCol As Integer, PrvW As Integer
Public MySource As String, RwSel As Boolean
Dim tm As Long
Dim MRC1Connect As Boolean, MRC2Connect As Boolean, UlasConnect As Boolean, TorilConnect As Boolean, Van2Connect As Boolean
Dim PanaboConnect As Boolean, BinosPolConnect As Boolean, YamahaConnect As Boolean, Van1Connect As Boolean, GensanConnect As Boolean
Dim YmarbelConnect As Boolean, StotomasConnect As Boolean, SamalConnect As Boolean, Van3Connect As Boolean
Dim WsaleConnect As Boolean, Whouse As Boolean, CbntianConnect As Boolean
Dim rsClass As New ADODB.Recordset


Private Sub cboClass_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboClass
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsClass.MoveFirst
            rsClass.Move i
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            'txtSuppCode.Text = ""
        End If
    End With

End Sub

Private Sub cboClass_Click()
    Call cboClass_Change
End Sub

Private Sub cboClass_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboClass.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboFind.SetFocus
End Sub


Private Sub cboFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdEvaluate_Click    '     cmdFind_Click
End Sub

Private Sub cbomos_Change()
    cbomos_Click
End Sub

Private Sub cbomos_Click()
Dim mos As Integer
    mos = Val(cboMos.Text)
    pmos = "-" & mos
    SDfrom.Value = DateAdd("m", pmos, Date)
   
    If cboMos.Text = "All" Then
        SDfrom.Value = SPordering.SalesDateFrom.Value
    Else
        flxSPH.TextMatrix(1, 75) = cboMos.Text & "mSOQ"
    End If
End Sub

Private Sub cbomos_DblClick()
    cbomos_Click
End Sub

Private Sub cbomos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 54 Or KeyAscii < 48 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsSupplier.MoveFirst
            rsSupplier.Move i
            txtSuppCode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            chkAllSupp.Value = 0
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppCode.Text = ""
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
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboFind.SetFocus
End Sub

Private Sub cboSupplier_LostFocus()
    If cboSupplier.Text = "" Then txtSuppCode.Text = ""
End Sub

Private Sub chkAllSupp_Click()
    If chkAllSupp.Value = 1 Then 'All Supplier
        cboSupplier.Text = ""
        txtSuppCode.Text = ""
    Else
        If txtSuppCode.Text = "" Then
            cboSupplier.Text = SPordering.cboSupplier.Text
            txtSuppCode.Text = SPordering.txtScode.Text
        End If
    End If
End Sub

Private Sub chkResize_Click()
    If chkResize.Value = 1 Then
        Timer1.Enabled = True
        tm = 0
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub DisableButtonOnProcess(Bln As Boolean)

    If Bln Then
        If ShowOpt = True Then
            cmdOptHide.Enabled = False
            FrameOpt.Enabled = False
        Else
            cmdOptShow.Enabled = False
        End If
        cmdEvaluate.Enabled = False
        cmdReplenish.Enabled = False
        cmdInsert.Enabled = False
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
        fScrol.Enabled = False
    Else
        If ShowOpt = True Then
            cmdOptHide.Enabled = True
            FrameOpt.Enabled = True
        Else
            cmdOptShow.Enabled = True
        End If
        cmdEvaluate.Enabled = True
        cmdReplenish.Enabled = True
        cmdInsert.Enabled = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        fScrol.Enabled = True
    End If
End Sub

Private Sub cmdEvaluate_Click()
    Dim rsSP As New ADODB.Recordset
    
    Dim MySP() As String, Mdl() As String, Qry As String
    Dim myW As String, myW1 As String, myW2 As String, myW3 As String
    Dim MySel As String, i As Long, r As Long, rM As Integer

    myW1 = ""
    If cboFind.Text <> "" Then
        MySP = Split(cboFind.Text, " ")
        For i = LBound(MySP) To UBound(MySP)
            If myW1 = "" Then
                'myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
                
                If updRef.Value = 0 Then
                    myW1 = "PARTNUMBER Like \'%" & MySP(i) & "%\'"
                Else
                    myW1 = "DESCRIPTION Like \'%" & MySP(i) & "%\'"
                End If
            Else
                'myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
                If updRef.Value = 0 Then
                    myW1 = myW1 & " AND PARTNUMBER Like \'%" & MySP(i) & "%\'"
                Else
                    myW1 = myW1 & " AND DESCRIPTION Like \'%" & MySP(i) & "%\'"
                End If
            End If
        Next i
        myW = myW1
    End If
    If txtSuppCode.Text <> "" And myW <> "" Then
        myW = myW & " AND SUPPLIERCODE = \'" & txtSuppCode.Text & "\'"
    ElseIf txtSuppCode.Text <> "" And myW = "" Then
        myW = "SUPPLIERCODE = '" & txtSuppCode.Text & "'"
    ElseIf txtSuppCode.Text = "" And myW = "" Then
        Exit Sub
    End If
    DoEvents
    tabProcessing.Visible = True
    Call ResizeCompleteHead
    Call CollAlign
    Call DisableButtonOnProcess(True)
    Qry = "CALL mrcwarehouse.sp_podata('" & myW & "', '" & Format(SDfrom.Value, "yyyy-mm-dd") & "', '" & Format(SDto.Value, "yyyy-mm-dd") & "');"
    Set rsSP = New ADODB.Recordset
    If rsSP.State = adStateOpen Then rsSP.Close
    With rsSP
        DoEvents
        .Open Qry, LYRcon, adOpenStatic
        If .RecordCount = 1 Then
            .MoveFirst
            myW2 = .Fields(0)
            myW3 = .Fields(1)
            If myW2 = "Error" Then
                MsgBox myW3, vbCritical + vbOKOnly, "DB ERROR"
                .Close
                flxSP.Clear
                flxSP.Rows = 1
                Exit Sub
            End If
        End If
        If .RecordCount > 0 Then
            flxSP.Rows = 1
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount + 1
            .MoveFirst
            For rM = 1 To .RecordCount
                flxSP.Rows = rM
                flxSP.TextMatrix(rM - 1, 1) = !PartNumber
                flxSP.TextMatrix(rM - 1, 2) = !Description
                flxSP.TextMatrix(rM - 1, 3) = !Ulas_QH
                flxSP.TextMatrix(rM - 1, 4) = !Ulas_BO
                flxSP.TextMatrix(rM - 1, 5) = !Ulas_SQ
                flxSP.TextMatrix(rM - 1, 6) = !Ulas_ST
                flxSP.TextMatrix(rM - 1, 11) = !Toril_QH
                flxSP.TextMatrix(rM - 1, 12) = !Toril_BO
                flxSP.TextMatrix(rM - 1, 13) = !Toril_SQ
                flxSP.TextMatrix(rM - 1, 14) = !Toril_ST
                flxSP.TextMatrix(rM - 1, 15) = !Van2_QH
                flxSP.TextMatrix(rM - 1, 16) = !Van2_BO
                flxSP.TextMatrix(rM - 1, 17) = !Van2_SQ
                flxSP.TextMatrix(rM - 1, 18) = !Van2_ST
                flxSP.TextMatrix(rM - 1, 19) = !Van56_QH
                flxSP.TextMatrix(rM - 1, 20) = !Van56_BO
                flxSP.TextMatrix(rM - 1, 21) = !Van56_SQ
                flxSP.TextMatrix(rM - 1, 22) = !Van56_ST
                flxSP.TextMatrix(rM - 1, 23) = !Panabo_QH
                flxSP.TextMatrix(rM - 1, 24) = !Panabo_BO
                flxSP.TextMatrix(rM - 1, 25) = !Panabo_SQ
                flxSP.TextMatrix(rM - 1, 26) = !Panabo_ST
                flxSP.TextMatrix(rM - 1, 31) = !Sob_QH
                flxSP.TextMatrix(rM - 1, 32) = !Sob_BO
                flxSP.TextMatrix(rM - 1, 33) = !Sob_SQ
                flxSP.TextMatrix(rM - 1, 34) = !Sob_ST
                flxSP.TextMatrix(rM - 1, 35) = !Gensan_QH
                flxSP.TextMatrix(rM - 1, 36) = !Gensan_BO
                flxSP.TextMatrix(rM - 1, 37) = !Gensan_SQ
                flxSP.TextMatrix(rM - 1, 38) = !Gensan_ST
                flxSP.TextMatrix(rM - 1, 39) = !Gante_QH
                flxSP.TextMatrix(rM - 1, 40) = !Gante_BO
                flxSP.TextMatrix(rM - 1, 41) = !Gante_SQ
                flxSP.TextMatrix(rM - 1, 42) = !Gante_ST
                flxSP.TextMatrix(rM - 1, 55) = !Van55_QH
                flxSP.TextMatrix(rM - 1, 56) = !Van55_BO
                flxSP.TextMatrix(rM - 1, 57) = !Van55_SQ
                flxSP.TextMatrix(rM - 1, 58) = !Van55_ST
                flxSP.TextMatrix(rM - 1, 59) = !WS_QH
                flxSP.TextMatrix(rM - 1, 60) = !WS_BO
                flxSP.TextMatrix(rM - 1, 61) = !WS_SQ
                flxSP.TextMatrix(rM - 1, 62) = !WS_ST
                flxSP.TextMatrix(rM - 1, 63) = !CAB_QH
                flxSP.TextMatrix(rM - 1, 64) = !CAB_BO
                flxSP.TextMatrix(rM - 1, 65) = !CAB_SQ
                flxSP.TextMatrix(rM - 1, 66) = !CAB_ST
                flxSP.TextMatrix(rM - 1, 71) = !WARE_QH
                flxSP.TextMatrix(rM - 1, 72) = !WARE_BO
                flxSP.TextMatrix(rM - 1, 73) = !WARE_ST 'NEW: IN-TRANSIT    old:6MSQ
                flxSP.TextMatrix(rM - 1, 77) = Format(!cost, "#,##0.00")
                flxSP.TextMatrix(rM - 1, 78) = IIf(IsNull(!ApplicModel), "", !ApplicModel)
                flxSP.TextMatrix(rM - 1, 79) = !SPAREPARTID
                flxSP.TextMatrix(rM - 1, 80) = IIf(IsNull(!DESCRIP2), "", !DESCRIP2)
                DoEvents
                pBar.Value = pBar.Value + 1
                .MoveNext
            Next rM
            
            flxSP.Visible = True
            If ShowOpt Then
                'flxSP.Top = 1920
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                'flxSP.Top = 1080
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
           ' Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
        Else
            flxSP.Clear
            flxSP.Rows = 1
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    Call AddInCombo
    pBar.Value = 0
    pBar.Max = flxSP.Rows
    For r = 0 To flxSP.Rows - 1
        pBar.Value = pBar.Value + 1
        Ulas6M = 0: Yamaha6M = 0: Toril6M = 0: Van16M = 0: Van26M = 0
        Mrc16m = 0: Panabo6M = 0: BinosPol6M = 0: GenSan6M = 0: Mrc26M = 0: Ymarbel6m = 0
        Stotomas6m = 0: Samal6m = 0: Van36m = 0: Wsale6m = 0: Cbntian6m = 0
        'Total Branch
        'SQ
        flxSP.TextMatrix(r, 67) = Val(flxSP.TextMatrix(r, 3)) + Val(flxSP.TextMatrix(r, 7)) + Val(flxSP.TextMatrix(r, 11)) + Val(flxSP.TextMatrix(r, 15)) + Val(flxSP.TextMatrix(r, 19)) + Val(flxSP.TextMatrix(r, 23)) + Val(flxSP.TextMatrix(r, 27)) + Val(flxSP.TextMatrix(r, 31)) + Val(flxSP.TextMatrix(r, 35)) + Val(flxSP.TextMatrix(r, 39)) + Val(flxSP.TextMatrix(r, 43)) + Val(flxSP.TextMatrix(r, 47)) + Val(flxSP.TextMatrix(r, 51)) + Val(flxSP.TextMatrix(r, 55)) + Val(flxSP.TextMatrix(r, 59)) + Val(flxSP.TextMatrix(r, 63))
        'BO
        flxSP.TextMatrix(r, 68) = Val(flxSP.TextMatrix(r, 4)) + Val(flxSP.TextMatrix(r, 8)) + Val(flxSP.TextMatrix(r, 12)) + Val(flxSP.TextMatrix(r, 16)) + Val(flxSP.TextMatrix(r, 20)) + Val(flxSP.TextMatrix(r, 24)) + Val(flxSP.TextMatrix(r, 28)) + Val(flxSP.TextMatrix(r, 32)) + Val(flxSP.TextMatrix(r, 36)) + Val(flxSP.TextMatrix(r, 40)) + Val(flxSP.TextMatrix(r, 44)) + Val(flxSP.TextMatrix(r, 48)) + Val(flxSP.TextMatrix(r, 52)) + Val(flxSP.TextMatrix(r, 56)) + Val(flxSP.TextMatrix(r, 60)) + Val(flxSP.TextMatrix(r, 64))
        'SQ
        flxSP.TextMatrix(r, 69) = Val(flxSP.TextMatrix(r, 5)) + Val(flxSP.TextMatrix(r, 9)) + Val(flxSP.TextMatrix(r, 13)) + Val(flxSP.TextMatrix(r, 17)) + Val(flxSP.TextMatrix(r, 21)) + Val(flxSP.TextMatrix(r, 25)) + Val(flxSP.TextMatrix(r, 29)) + Val(flxSP.TextMatrix(r, 33)) + Val(flxSP.TextMatrix(r, 37)) + Val(flxSP.TextMatrix(r, 41)) + Val(flxSP.TextMatrix(r, 45)) + Val(flxSP.TextMatrix(r, 49)) + Val(flxSP.TextMatrix(r, 53)) + Val(flxSP.TextMatrix(r, 57)) + Val(flxSP.TextMatrix(r, 61)) + Val(flxSP.TextMatrix(r, 65))
        'ST0ck transfer
        flxSP.TextMatrix(r, 70) = Val(flxSP.TextMatrix(r, 6)) + Val(flxSP.TextMatrix(r, 10)) + Val(flxSP.TextMatrix(r, 14)) + Val(flxSP.TextMatrix(r, 18)) + Val(flxSP.TextMatrix(r, 22)) + Val(flxSP.TextMatrix(r, 26)) + Val(flxSP.TextMatrix(r, 30)) + Val(flxSP.TextMatrix(r, 34)) + Val(flxSP.TextMatrix(r, 38)) + Val(flxSP.TextMatrix(r, 42)) + Val(flxSP.TextMatrix(r, 46)) + Val(flxSP.TextMatrix(r, 50)) + Val(flxSP.TextMatrix(r, 54)) + Val(flxSP.TextMatrix(r, 58)) + Val(flxSP.TextMatrix(r, 62)) + Val(flxSP.TextMatrix(r, 66))
        'Total Qty OnHand (TTLQH)
        flxSP.TextMatrix(r, 74) = Val(flxSP.TextMatrix(r, 67)) + Val(flxSP.TextMatrix(r, 71))
                'Branch Total Sales Qty(69) Less TtlQH(74)+Warehouse InTransit(73) + Warehouse BO(72)
        spSOQ = Val(flxSP.TextMatrix(r, 69)) - (Val(flxSP.TextMatrix(r, 74)) + Val(flxSP.TextMatrix(r, 73)) + Val(flxSP.TextMatrix(r, 72)))
        flxSP.TextMatrix(r, 75) = IIf(spSOQ < 0, 0, Int(spSOQ + 0.49))

'        flxSP.TextMatrix(r, 74) = Ulas6M + Yamaha6M + Toril6M + Van16M + Van26M + Panabo6M + BinosPol6M + Mrc16m + GenSan6M + Mrc26M + Ymarbel6m + Stotomas6m + Samal6m + Van36m + Wsale6m + Cbntian6m + Val(flxSP.TextMatrix(r, 73))
'
'        'suggested order Qty.  = 6m Sales Gante+6m Sales Qty. Branch -( Branch QH + Main QH) -  Back Order Main only
'        'Before Jan. 4, 2013
'        'spSOQ = Val(flxSP.TextMatrix(r, 42)) - (Val(flxSP.TextMatrix(r, 35)) + Val(flxSP.TextMatrix(r, 39))) - Val(flxSP.TextMatrix(r, 40))
'        'new Monthly Avg (6Mos) * Order Mos  - (QH+BO)
'        spSOQ = ((Val(flxSP.TextMatrix(r, 74) / 6)) * Val(cbomos.Text)) - (Val(flxSP.TextMatrix(r, 67)) + Val(flxSP.TextMatrix(r, 71)) + Val(flxSP.TextMatrix(r, 72)))
'        flxSP.TextMatrix(r, 75) = IIf(spSOQ < 0, 0, Int(spSOQ + 0.49))
        
    Next r
    pBar.Visible = False
    flxSP.Col = 71: flxSP.Row = 0: flxSP.ColSel = 73: flxSP.RowSel = flxSP.Rows - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(150, 230, 150)
    If flxSP.Rows >= 1 Then
        'If flxSP.TextMatrix(1, 3) <> "" Then
                         'QTY on Hand
        If flxSP.TextMatrix(0, 67) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
    tabProcessing.Visible = False
    Call DisableButtonOnProcess(False)

End Sub

Private Sub cmdFind_Click()
    Dim rsSP As New ADODB.Recordset
   
    Dim MySP() As String, Mdl() As String, Qry As String
    Dim myW As String, myW1 As String, myW2 As String, myW3 As String
    Dim MySel As String, i As Long, r As Long, rM As Integer
    
    Dim spSOQ As Double
    Dim Ulas6M      As Double
    Dim Yamaha6M    As Double
    Dim Toril6M     As Double
    Dim Van16M      As Double
    Dim Van26M      As Double
    Dim Mrc16m      As Double
    Dim Panabo6M    As Double
    Dim BinosPol6M  As Double
    Dim GenSan6M    As Double
    Dim Mrc26M      As Double
    Dim Ymarbel6m   As Double
    Dim Stotomas6m  As Double
    Dim Samal6m     As Double
    Dim Van36m      As Double
    Dim Wsale6m     As Double
    Dim Cbntian6m   As Double
    
    If cboFind.Text = "" Then Exit Sub
    MySP = Split(cboFind.Text, " ")
    myW1 = ""
    For i = LBound(MySP) To UBound(MySP)
        If myW1 = "" Then
            'myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            
            If updRef.Value = 0 Then
                myW1 = "(SPAREPARTS.PARTNUMBER) Like '%" & MySP(i) & "%'"
            Else
                myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            End If
        Else
            'myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            If updRef.Value = 0 Then
                myW1 = myW1 & " AND (SPAREPARTS.PARTNUMBER) Like '%" & MySP(i) & "%'"
            Else
                myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            End If
        End If
    Next i
    myW = myW1
    If txtSuppCode.Text <> "" Then
        myW = "(" & myW & ") AND ((SPAREPARTS.SUPPLIERCODE) = '" & txtSuppCode.Text & "')"
    End If
    Call ResizeCompleteHead
    Call CollAlign
    MySel = "SELECT  SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT From SPAREPARTS WHERE (" & myW & ")  ORDER BY SPAREPARTS.PARTNUMBER, SPAREPARTS.QTYONHAND DESC;"
    With rsSP
        .Open MySel, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 1
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount * 2
            .MoveFirst
            For rM = 1 To .RecordCount
                    flxSP.Rows = rM
                    flxSP.TextMatrix(rM - 1, 1) = !PartNumber
                    flxSP.TextMatrix(rM - 1, 2) = !Description
                    flxSP.TextMatrix(rM - 1, 71) = !QtyOnHand
                    flxSP.TextMatrix(rM - 1, 72) = IIf(IsNull(!BACKORDER), "", !BACKORDER)
                    flxSP.TextMatrix(rM - 1, 77) = Format(!cost, "#,##0.00")
                    flxSP.TextMatrix(rM - 1, 79) = !SPAREPARTID
                    flxSP.TextMatrix(rM - 1, 78) = IIf(IsNull(!ApplicModel), "", !ApplicModel)
                    flxSP.TextMatrix(rM - 1, 80) = IIf(IsNull(!MyStat), "", !MyStat)
                    pBar.Value = pBar.Value + 1
                .MoveNext
            Next rM
            
            flxSP.Visible = True
            If ShowOpt Then
                'flxSP.Top = 1920
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                'flxSP.Top = 1080
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
           ' Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
        Else
            flxSP.Clear
            flxSP.Rows = 1
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    With flxSP
        
        For i = .Rows - 1 To 0 Step -1
            'MRC WAREHOUSE
            .TextMatrix(i, 73) = GetSalesQty(Val(.TextMatrix(i, 79)))
            
            '---------------------------------------------------------
            '** Stock Transfer
            'Moto-r Ulas
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 28)) & " AND SPITR.SOURCE='MOTO-R ULAS';"
            '.TextMatrix(i, 22) = Val(GetFieldValue(Qry))
            'Yamaha Motorhub
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 28)) & " AND SPITR.SOURCE='YAMAHA MOTORHUB';"
            '.TextMatrix(i, 23) = Val(GetFieldValue(Qry))
            
            'MC Pasrts Depot
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 28)) & " AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            '.TextMatrix(i, 24) = Val(GetFieldValue(Qry))
            '---------------------------------------------------------
            
            
            pBar.Value = IIf(pBar.Value = pBar.Max, pBar.Max, pBar.Value + 1)
            DoEvents
            If InStr(.TextMatrix(i, 80), "INACTIVE") > 0 Then
                .Col = 1: .Row = i: .ColSel = 78: .RowSel = i: .FillStyle = flexFillRepeat: .CellForeColor = RGB(255, 0, 0)
            End If
        Next i
        
    End With
    
    
    'Pbar.Visible = False
    Call AddInCombo
    pBar.Value = 0
    pBar.Max = flxSP.Rows
    For r = 0 To flxSP.Rows - 1
        pBar.Value = pBar.Value + 1
        Ulas6M = 0: Yamaha6M = 0: Toril6M = 0: Van16M = 0: Van26M = 0
        Mrc16m = 0: Panabo6M = 0: BinosPol6M = 0: GenSan6M = 0: Mrc26M = 0: Ymarbel6m = 0
        Stotomas6m = 0: Samal6m = 0: Van36m = 0: Wsale6m = 0: Cbntian6m = 0
        
        
        '**Query Branch SKU
        'Moto-r Ulas
        If UlasConnect Then
            Qry = ""
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 3) = Val(GetFieldValueBranch(Qry, "ULAS"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 4) = Val(GetFieldValueBranch(Qry, "ULAS"))
            flxSP.TextMatrix(r, 5) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "ULAS")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"      'AND SPITR.SOURCE='MOTO-R ULAS';"
            flxSP.TextMatrix(r, 6) = Val(GetFieldValueBranch(Qry, "ULAS"))
            Ulas6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "ULAS")
        End If
        'Yamaha MotorHub
        If YamahaConnect Then
            Qry = ""
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 7) = Val(GetFieldValueBranch(Qry, "YAMAHA"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 8) = Val(GetFieldValueBranch(Qry, "YAMAHA"))
            flxSP.TextMatrix(r, 9) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "YAMAHA")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='YAMAHA MOTORHUB';"
            flxSP.TextMatrix(r, 10) = Val(GetFieldValueBranch(Qry, "YAMAHA"))
            Yamaha6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "YAMAHA")
        End If
        
        'MC Parts Depot Corporation
        If TorilConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 11) = Val(GetFieldValueBranch(Qry, "TORIL"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 12) = Val(GetFieldValueBranch(Qry, "TORIL"))
            flxSP.TextMatrix(r, 13) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "TORIL")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 14) = Val(GetFieldValueBranch(Qry, "TORIL"))
            Toril6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "TORIL")
        End If
       
        'Van Selling I
        If Van1Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 15) = Val(GetFieldValueBranch(Qry, "VAN1"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 16) = Val(GetFieldValueBranch(Qry, "VAN1"))
            flxSP.TextMatrix(r, 17) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "VAN1")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 45)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 18) = Val(GetFieldValueBranch(Qry, "VAN1"))
            Van16M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "VAN1")
        End If
        
        'Van Selling II
        If Van2Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 19) = Val(GetFieldValueBranch(Qry, "VAN2"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 20) = Val(GetFieldValueBranch(Qry, "VAN2"))
            flxSP.TextMatrix(r, 21) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "VAN2")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 22) = Val(GetFieldValueBranch(Qry, "VAN2"))
            Van26M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "VAN2")
        End If
        
        'MRC PANABO
        If PanaboConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 23) = Val(GetFieldValueBranch(Qry, "PANABO"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 24) = Val(GetFieldValueBranch(Qry, "PANABO"))
            flxSP.TextMatrix(r, 25) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "PANABO")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 26) = Val(GetFieldValueBranch(Qry, "PANABO"))
            Panabo6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "PANABO")
        End If
        
        'BINOS POLOMOLOK
        If BinosPolConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 27) = Val(GetFieldValueBranch(Qry, "POLOMOLOK"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 28) = Val(GetFieldValueBranch(Qry, "POLOMOLOK"))
            flxSP.TextMatrix(r, 29) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "POLOMOLOK")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 30) = Val(GetFieldValueBranch(Qry, "POLOMOLOK"))
            BinosPol6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "POLOMOLOK")
        End If
        
        'MRC1
        If MRC1Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 31) = Val(GetFieldValueBranch(Qry, "MRC1"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 32) = Val(GetFieldValueBranch(Qry, "MRC1"))
            flxSP.TextMatrix(r, 33) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "MRC1")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 34) = Val(GetFieldValueBranch(Qry, "MRC1"))
            Mrc16m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "MRC1")
        End If
        
        'GENSAN
        If GensanConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 35) = Val(GetFieldValueBranch(Qry, "GENSAN"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 36) = Val(GetFieldValueBranch(Qry, "GENSAN"))
            flxSP.TextMatrix(r, 37) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "GENSAN")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 38) = Val(GetFieldValueBranch(Qry, "GENSAN"))
            GenSan6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "GENSAN")
        End If
        
        'MRC 2 GANTE
        If MRC2Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 39) = Val(GetFieldValueBranch(Qry, "GANTE"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 40) = Val(GetFieldValueBranch(Qry, "GANTE"))
            flxSP.TextMatrix(r, 41) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "GANTE")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 42) = Val(GetFieldValueBranch(Qry, "GANTE"))
            Mrc26M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "GANTE")
        End If
        
        If YmarbelConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 43) = Val(GetFieldValueBranch(Qry, "MARBEL"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 44) = Val(GetFieldValueBranch(Qry, "MARBEL"))
            flxSP.TextMatrix(r, 45) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "MARBEL")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 46) = Val(GetFieldValueBranch(Qry, "MARBEL"))
            Ymarbel6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "MARBEL")
        End If
        
        If StotomasConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 47) = Val(GetFieldValueBranch(Qry, "STOTOMAS"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 48) = Val(GetFieldValueBranch(Qry, "STOTOMAS"))
            flxSP.TextMatrix(r, 49) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "STOTOMAS")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 50) = Val(GetFieldValueBranch(Qry, "STOTOMAS"))
            Stotomas6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "STOTOMAS")
        End If
        
        If SamalConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 51) = Val(GetFieldValueBranch(Qry, "SAMAL"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 52) = Val(GetFieldValueBranch(Qry, "SAMAL"))
            flxSP.TextMatrix(r, 53) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "SAMAL")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 54) = Val(GetFieldValueBranch(Qry, "SAMAL"))
            Samal6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "SAMAL")
        End If
        
        If Van3Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 55) = Val(GetFieldValueBranch(Qry, "VAN3"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 56) = Val(GetFieldValueBranch(Qry, "VAN3"))
            flxSP.TextMatrix(r, 57) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "VAN3")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 58) = Val(GetFieldValueBranch(Qry, "VAN3"))
            Van36m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "VAN3")
        End If
        
        If WsaleConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 59) = Val(GetFieldValueBranch(Qry, "WSALE"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 60) = Val(GetFieldValueBranch(Qry, "WSALE"))
            flxSP.TextMatrix(r, 61) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "WSALE")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 62) = Val(GetFieldValueBranch(Qry, "WSALE"))
            Wsale6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "WSALE")
        End If
        
        If CbntianConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 63) = Val(GetFieldValueBranch(Qry, "CABANTIAN"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 64) = Val(GetFieldValueBranch(Qry, "CABANTIAN"))
            flxSP.TextMatrix(r, 65) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "CABANTIAN")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 66) = Val(GetFieldValueBranch(Qry, "CABANTIAN"))
            Wsale6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "CABANTIAN")
        End If
        
        
        'Total Branch
        'SQ
        flxSP.TextMatrix(r, 67) = Val(flxSP.TextMatrix(r, 3)) + Val(flxSP.TextMatrix(r, 7)) + Val(flxSP.TextMatrix(r, 11)) + Val(flxSP.TextMatrix(r, 15)) + Val(flxSP.TextMatrix(r, 19)) + Val(flxSP.TextMatrix(r, 23)) + Val(flxSP.TextMatrix(r, 27)) + Val(flxSP.TextMatrix(r, 31)) + Val(flxSP.TextMatrix(r, 35)) + Val(flxSP.TextMatrix(r, 39)) + Val(flxSP.TextMatrix(r, 43)) + Val(flxSP.TextMatrix(r, 47)) + Val(flxSP.TextMatrix(r, 51)) + Val(flxSP.TextMatrix(r, 55)) + Val(flxSP.TextMatrix(r, 59))
        'BO
        flxSP.TextMatrix(r, 68) = Val(flxSP.TextMatrix(r, 4)) + Val(flxSP.TextMatrix(r, 8)) + Val(flxSP.TextMatrix(r, 12)) + Val(flxSP.TextMatrix(r, 16)) + Val(flxSP.TextMatrix(r, 20)) + Val(flxSP.TextMatrix(r, 24)) + Val(flxSP.TextMatrix(r, 28)) + Val(flxSP.TextMatrix(r, 32)) + Val(flxSP.TextMatrix(r, 36)) + Val(flxSP.TextMatrix(r, 40)) + Val(flxSP.TextMatrix(r, 44)) + Val(flxSP.TextMatrix(r, 48)) + Val(flxSP.TextMatrix(r, 52)) + Val(flxSP.TextMatrix(r, 56)) + Val(flxSP.TextMatrix(r, 60))
        'SQ
        flxSP.TextMatrix(r, 69) = Val(flxSP.TextMatrix(r, 5)) + Val(flxSP.TextMatrix(r, 9)) + Val(flxSP.TextMatrix(r, 13)) + Val(flxSP.TextMatrix(r, 17)) + Val(flxSP.TextMatrix(r, 21)) + Val(flxSP.TextMatrix(r, 25)) + Val(flxSP.TextMatrix(r, 29)) + Val(flxSP.TextMatrix(r, 33)) + Val(flxSP.TextMatrix(r, 37)) + Val(flxSP.TextMatrix(r, 41)) + Val(flxSP.TextMatrix(r, 45)) + Val(flxSP.TextMatrix(r, 49)) + Val(flxSP.TextMatrix(r, 53)) + Val(flxSP.TextMatrix(r, 57)) + Val(flxSP.TextMatrix(r, 61))
        'ST0ck transfer
        flxSP.TextMatrix(r, 70) = Val(flxSP.TextMatrix(r, 6)) + Val(flxSP.TextMatrix(r, 10)) + Val(flxSP.TextMatrix(r, 14)) + Val(flxSP.TextMatrix(r, 18)) + Val(flxSP.TextMatrix(r, 22)) + Val(flxSP.TextMatrix(r, 26)) + Val(flxSP.TextMatrix(r, 30)) + Val(flxSP.TextMatrix(r, 34)) + Val(flxSP.TextMatrix(r, 38)) + Val(flxSP.TextMatrix(r, 42)) + Val(flxSP.TextMatrix(r, 46)) + Val(flxSP.TextMatrix(r, 50)) + Val(flxSP.TextMatrix(r, 54)) + Val(flxSP.TextMatrix(r, 58)) + Val(flxSP.TextMatrix(r, 62))
        
        flxSP.TextMatrix(r, 74) = Ulas6M + Yamaha6M + Toril6M + Van16M + Van26M + Panabo6M + BinosPol6M + Mrc16m + GenSan6M + Mrc26M + Ymarbel6m + Stotomas6m + Samal6m + Van36m + Wsale6m + Cbntian6m + Val(flxSP.TextMatrix(r, 73))
        
        'suggested order Qty.  = 6m Sales Gante+6m Sales Qty. Branch -( Branch QH + Main QH) -  Back Order Main only
        'Before Jan. 4, 2013
        'spSOQ = Val(flxSP.TextMatrix(r, 42)) - (Val(flxSP.TextMatrix(r, 35)) + Val(flxSP.TextMatrix(r, 39))) - Val(flxSP.TextMatrix(r, 40))
        'new Monthly Avg (6Mos) * Order Mos  - (QH+BO)
        spSOQ = ((Val(flxSP.TextMatrix(r, 74) / 6)) * Val(cboMos.Text)) - (Val(flxSP.TextMatrix(r, 67)) + Val(flxSP.TextMatrix(r, 71)) + Val(flxSP.TextMatrix(r, 72)))
        flxSP.TextMatrix(r, 75) = IIf(spSOQ < 0, 0, Int(spSOQ + 0.49))
        
    Next r
    pBar.Visible = False
    flxSP.Col = 71: flxSP.Row = 0: flxSP.ColSel = 73: flxSP.RowSel = flxSP.Rows - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(150, 230, 150)
    If flxSP.Rows >= 1 Then
        'If flxSP.TextMatrix(1, 3) <> "" Then
                         'QTY on Hand
        If flxSP.TextMatrix(0, 67) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
    
    
 End Sub

Function GetSalesQty(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    'old Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & SDto.Value & "));"
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)<>'TRANSFER') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
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
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

Function GetSalesQty6mos(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
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
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty6mos = MySales - MyRet
End Function

Function GetSalesQty6MBranch(MyID As Long, myCON As String) As Long
    Dim rsQty As New ADODB.Recordset
    Dim MyQ As String
    Dim MySales As Long, MyRet As Long
    
    MyQ = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        If myCON = "VAN1" Then
            .Open MyQ, conVAN1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN2" Then
            .Open MyQ, conVAN2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC2" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MRC1" Then
            .Open MyQ, conMRC1, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "ULAS" Then
            .Open MyQ, conUlas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "TORIL" Then
            .Open MyQ, conToril, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "YAMAHA" Then
            .Open MyQ, conYM, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "PANABO" Then
            .Open MyQ, conPanabo, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "POLOMOLOK" Then
            .Open MyQ, conPol, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GENSAN" Then
            .Open MyQ, conGenSan, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "GANTE" Then
            .Open MyQ, conMRC2, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "MARBEL" Then
            .Open MyQ, conYmarbel, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "STOTOMAS" Then
            .Open MyQ, conStotomas, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "SAMAL" Then
            .Open MyQ, conSamal, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "VAN3" Then
            .Open MyQ, conVan3, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "WSALE" Then
            .Open MyQ, conWsale, adOpenKeyset, adLockOptimistic
        ElseIf myCON = "CABANTIAN" Then
            .Open MyQ, conCabantian, adOpenKeyset, adLockOptimistic
        End If
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    MyQ = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueBranch(MyQ, myCON))
    GetSalesQty6MBranch = MySales - MyRet
End Function


Function GetSalesQty6mosMU(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConMU, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMU(Qry))
    GetSalesQty6mosMU = MySales - MyRet
End Function

Function GetSalesQty6mosYM(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConYM, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueYM(Qry))
    GetSalesQty6mosYM = MySales - MyRet
End Function

Function GetSalesQty6mosMD(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConMD, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMD(Qry))
    GetSalesQty6mosMD = MySales - MyRet
End Function

Function GetSalesQty6mosVAN(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConVAN, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueVAN(Qry))
    GetSalesQty6mosVAN = MySales - MyRet
End Function


Function GetSalesQty6mosMRC2(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrnConMRC2, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMRC2(Qry))
    GetSalesQty6mosMRC2 = MySales - MyRet
End Function



Function GetSalesQtyMain(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
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
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQtyMain = MySales - MyRet
End Function
Private Sub AddInCombo()
    Dim X As Integer, W As Boolean
    
    If cboFind.Text = "" Then Exit Sub
    W = False
    For X = 0 To cboFind.ListCount - 1
        If cboFind.Text = cboFind.List(X) Then
            If X = 0 Then
                W = True
                Exit For
            Else
                cboFind.RemoveItem X
                Exit For
            End If
        End If
    Next X
    If W = False Then cboFind.AddItem cboFind.Text, 0
End Sub

Private Sub cmdInsert_Click()
    Dim i As Integer
    
    
    If MySource = "SPORDERING" Then
        With SPordering.flxParts
            If .TextMatrix(.Rows - 1, 1) = "" Then  'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem .Rows - 1
                End If
            End If
            For i = 0 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 1)) Then
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & .TextMatrix(i, 13)
                        'Add applicable model
                        '                        1-PARTNUMBER                     2-DESCRIPTION                    3-ORDER                          4-COST                                   6-TOTAL                                                                          13-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 10)
                        '                        1-PARTNUMBER                     2-DESCRIPTION                    3-ORDER                          4-COST                       6-TOTAL                   10-spid                                                                        12-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 11) & vbTab & vbTab & flxSP.TextMatrix(i, 10) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 12)
                        'Design for  BrANCH     1-PARTNUMBER                       2-DESCRIPTION                  3-ORDER                          4-cost                       5-Total                    6-Spid                                                                        12-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 12) & vbTab & flxSP.TextMatrix(i, 14) & vbTab & vbTab & flxSP.TextMatrix(i, 13) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 15)
                        'Design for  Main     1-PARTNUMBER                       2-DESCRIPTION                  3-ORDER                          4-cost                       5-Total                    6-Spid                                                                        12-APPLICABLE MODEL
                        .AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 76) & vbTab & flxSP.TextMatrix(i, 77) & vbTab & vbTab & flxSP.TextMatrix(i, 79) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 78)
             
                    End If
                End If
            Next i
            If .Rows = 1 Then .AddItem "1"
        End With
    End If
    If MySource = "RECEIVING" Then
        'Unused code because we have separated program for receiving
        'ith SPincoming.flxOrd
        '    If .TextMatrix(.Rows - 1, 7) = "" And .TextMatrix(.Rows - 1, 2) = "" Then 'Remove last row if empty
                
        '        If .Rows = 2 Then
        '            .Rows = 1
        '        Else
        '             .RemoveItem .Rows - 1
        '        End If
        '    End If
        '    For i = 1 To flxSP.Rows - 1
        '        If flxSP.TextMatrix(i, 0) <> "" Then
        '            If Not SPinFlx(flxSP.TextMatrix(i, 9)) Then
                        '.AddItem vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 1)
                        '
        '                .AddItem vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 1)
        '
        '            End If
        '        End If
        '    Next i
        '    If .Rows = 1 Then .Rows = 2
        'End With
    End If
End Sub
Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    If MySource = "SPORDERING" Then
        With SPordering.flxParts
            For i = 1 To .Rows - 1
                If myPN = .TextMatrix(i, 1) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
    'Unused code for receiving
    'If MySource = "RECEIVING" Then
    '    With SPincoming.flxOrd
    '        For i = 1 To .Rows - 1
    '            If MyID = .TextMatrix(i, 7) Then
    '                SPinFlx = True
    '                Exit Function
    '            End If
    '        Next i
    '    End With
    '    SPinFlx = False
    '    Exit Function
    'End If
End Function

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Call SPordering.ComputeTotal
    Unload Me
End Sub

Private Sub cmdOptHide_Click()
    FrameOpt.Visible = False
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = False
    cmdOptShow.Visible = True
    flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
    flxSP.Top = FrameOpt.Height + flxSPH.Height
    cmdInsert.Top = flxSP.Top + flxSP.Height + 120
    cmdOK.Top = flxSP.Top + flxSP.Height + 120
    cmdCancel.Top = flxSP.Top + flxSP.Height + 120
    pBar.Top = flxSP.Top + flxSP.Height
    Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 880 + Frame1.Height
    Frame1.Top = cmdOK.Top + cmdOK.Height + 100
    fmAnotation.Top = Frame1.Top + Frame1.Height + 120
    fScrol.Top = flxSPH.Top + 15
    fScrol.Height = flxSPH.RowHeight(0)
    fScrol.Left = flxSPH.ColPos(6) + 80
    
    'txtModel.Text = ""
    ShowOpt = False
End Sub

Private Sub cmdOptShow_Click()
    FrameOpt.Visible = True
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = True
    cmdOptShow.Visible = False
    flxSP.Visible = True
    flxSPH.Top = FrameOpt.Top + FrameOpt.Height
    flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
    cmdInsert.Top = flxSP.Top + flxSP.Height + 120
    cmdOK.Top = flxSP.Top + flxSP.Height + 120
    cmdCancel.Top = flxSP.Top + flxSP.Height + 120
    pBar.Top = flxSP.Top + flxSP.Height
'    Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
'    fmAnotation.Top = cmdOK.Top + cmdOK.Height + 900 - fmAnotation.Height + 120
    Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 880 + Frame1.Height
    Frame1.Top = cmdOK.Top + cmdOK.Height + 100
    fmAnotation.Top = Frame1.Top + Frame1.Height + 120
    
    fScrol.Top = flxSPH.Top + 15
    fScrol.Height = flxSPH.RowHeight(0)
    fScrol.Left = flxSPH.ColPos(6) + 80
    fScrol.ZOrder
    ShowOpt = True
End Sub
Private Sub cmdReplenish_Click()
    Dim rsRpnsh As New ADODB.Recordset
    Dim Qry As String, r As Long, spSOQ As Double, rM As Integer
    
    Dim Ulas6M      As Double
    Dim Yamaha6M    As Double
    Dim Toril6M     As Double
    Dim Van16M      As Double
    Dim Van26M      As Double
    Dim Mrc16m      As Double
    Dim Panabo6M    As Double
    Dim BinosPol6M  As Double
    Dim GenSan6M    As Double
    Dim Mrc26M      As Double
    Dim Ymarbel6m   As Double
    Dim Stotomas6m  As Double
    Dim Samal6m     As Double
    Dim Van36m      As Double
    Dim Wsale6m     As Double
    
    cmdReplenish.Enabled = False
    cmdInsert.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    Call ResizeCompleteHead
    Call CollAlign
    If txtSuppCode.Text = "" Then
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND SPAREPARTS.DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY SPAREPARTS.PARTNUMBER;"
    Else
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND SPAREPARTS.SUPPLIERCODE='" & txtSuppCode.Text & "' AND SPAREPARTS.DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY SPAREPARTS.PARTNUMBER;"
    End If
    
    With rsRpnsh
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 0
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount * 2
            .MoveFirst
            For rM = 1 To .RecordCount                              'LOAD SPAREPARTS INFO OF MOTO-R GANTE
                    flxSP.Rows = rM
                    flxSP.TextMatrix(rM - 1, 1) = !PartNumber
                    flxSP.TextMatrix(rM - 1, 2) = !Description
                    flxSP.TextMatrix(rM - 1, 71) = !QtyOnHand
                    flxSP.TextMatrix(rM - 1, 72) = IIf(IsNull(!BACKORDER), "", !BACKORDER)
                    flxSP.TextMatrix(rM - 1, 77) = Format(!cost, "#,##0.00")
                    flxSP.TextMatrix(rM - 1, 78) = IIf(IsNull(!ApplicModel), "", !ApplicModel)
                    flxSP.TextMatrix(rM - 1, 79) = !SPAREPARTID
                    flxSP.TextMatrix(rM - 1, 80) = IIf(IsNull(!MyStat), "", !MyStat)
                    pBar.Value = pBar.Value + 1
                .MoveNext
            Next rM
            
            flxSP.Visible = True
            If ShowOpt Then
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
            'Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
           
        Else
            flxSP.Clear
            flxSP.Rows = 1
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    
   
    For i = flxSP.Rows - 1 To 0 Step -1
        'Main Branch
        flxSP.TextMatrix(i, 73) = GetSalesQty(Val(flxSP.TextMatrix(i, 79)))
        
        
        '----------------------------------------------------------
        '** Stock Transfer - Order Taking
        'Moto-r Ulas
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(i, 28)) & " AND SPITR.SOURCE='MOTO-R ULAS';"
        'flxSP.TextMatrix(i, 22) = Val(GetFieldValue(Qry))
        'Yamaha Motorhub
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(i, 28)) & " AND SPITR.SOURCE='YAMAHA MOTORHUB';"
        'flxSP.TextMatrix(i, 23) = Val(GetFieldValue(Qry))
        
        'MC Pasrts Depot
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(i, 28)) & " AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
        'flxSP.TextMatrix(i, 24) = Val(GetFieldValue(Qry))
        '-----------------------------------------------------------
        
        
        pBar.Value = IIf(pBar.Value = pBar.Max, pBar.Max, pBar.Value + 1)
        DoEvents
        If flxSP.TextMatrix(i, 80) = "INACTIVE" Then
            flxSP.Col = 1: flxSP.Row = i: flxSP.ColSel = 79: flxSP.RowSel = i: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(255, 0, 0)
        End If
        '***Option Sales Quantity
        Select Case OptrSQ.Value
            Case 0
                If Val(flxSP.TextMatrix(i, 73)) <> Val(txtSQ.Text) Then 'flxSP.TextMatrix(i, 9) = "x"
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 1
                If Val(flxSP.TextMatrix(i, 73)) <= Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 2
                If Val(flxSP.TextMatrix(i, 73)) < Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 3
                If Val(flxSP.TextMatrix(i, 73)) >= Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 4
                If Val(flxSP.TextMatrix(i, 73)) > Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 5
                If Val(flxSP.TextMatrix(i, 73)) = Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
        End Select
    Next i
    Call AddInCombo
    'Pbar.Visible = False
    'Get 6 mos Sales quantity
    
    'pBar.Max = flxSP.Rows - 1
    'For r = 0 To flxSP.Rows - 1
    pBar.Value = 0
    pBar.Max = flxSP.Rows
    For r = 0 To flxSP.Rows - 1
        pBar.Value = pBar.Value + 1
        Ulas6M = 0: Yamaha6M = 0: Toril6M = 0: Van16M = 0: Van26M = 0
        Mrc16m = 0: Panabo6M = 0: BinosPol6M = 0: GenSan6M = 0: Mrc26M = 0: Ymarbel6m = 0
        Stotomas6m = 0: Samal6m = 0: Van36m = 0: Wsale6m = 0: Cbntian6m = 0
        '**Query Branch SKU
        'Moto-r Ulas
        If UlasConnect Then
            Qry = ""
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 3) = Val(GetFieldValueBranch(Qry, "ULAS"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 4) = Val(GetFieldValueBranch(Qry, "ULAS"))
            flxSP.TextMatrix(r, 5) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "ULAS")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"      'AND SPITR.SOURCE='MOTO-R ULAS';"
            flxSP.TextMatrix(r, 6) = Val(GetFieldValueBranch(Qry, "ULAS"))
            Ulas6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "ULAS")
        End If
        'Yamaha MotorHub
        If YamahaConnect Then
            Qry = ""
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 7) = Val(GetFieldValueBranch(Qry, "YAMAHA"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 8) = Val(GetFieldValueBranch(Qry, "YAMAHA"))
            flxSP.TextMatrix(r, 9) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "YAMAHA")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='YAMAHA MOTORHUB';"
            flxSP.TextMatrix(r, 10) = Val(GetFieldValueBranch(Qry, "YAMAHA"))
            Yamaha6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "YAMAHA")
        End If
        
        'MC Parts Depot Corporation
        If TorilConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 11) = Val(GetFieldValueBranch(Qry, "TORIL"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 12) = Val(GetFieldValueBranch(Qry, "TORIL"))
            flxSP.TextMatrix(r, 13) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "TORIL")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 14) = Val(GetFieldValueBranch(Qry, "TORIL"))
            Toril6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "TORIL")
        End If
       
        'Van Selling I
        If Van1Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 15) = Val(GetFieldValueBranch(Qry, "VAN1"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 16) = Val(GetFieldValueBranch(Qry, "VAN1"))
            flxSP.TextMatrix(r, 17) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "VAN1")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 18) = Val(GetFieldValueBranch(Qry, "VAN1"))
            Van16M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "VAN1")
        End If
        'Van Selling II
        If Van2Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 19) = Val(GetFieldValueBranch(Qry, "VAN2"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 20) = Val(GetFieldValueBranch(Qry, "VAN2"))
            flxSP.TextMatrix(r, 21) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "VAN2")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 22) = Val(GetFieldValueBranch(Qry, "VAN2"))
            Van26M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "VAN2")
        End If
        
        'MRC PANABO
        If PanaboConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 23) = Val(GetFieldValueBranch(Qry, "PANABO"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 24) = Val(GetFieldValueBranch(Qry, "PANABO"))
            flxSP.TextMatrix(r, 25) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "PANABO")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 26) = Val(GetFieldValueBranch(Qry, "PANABO"))
            Panabo6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "PANABO")
        End If
        
        'BINOS POLOMOLOK
        If BinosPolConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 27) = Val(GetFieldValueBranch(Qry, "POLOMOLOK"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 28) = Val(GetFieldValueBranch(Qry, "POLOMOLOK"))
            flxSP.TextMatrix(r, 29) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "POLOMOLOK")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 30) = Val(GetFieldValueBranch(Qry, "POLOMOLOK"))
            BinosPol6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "POLOMOLOK")
        End If
        
        'MRC1
        If MRC1Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 31) = Val(GetFieldValueBranch(Qry, "MRC1"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 32) = Val(GetFieldValueBranch(Qry, "MRC1"))
            flxSP.TextMatrix(r, 33) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "MRC1")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 34) = Val(GetFieldValueBranch(Qry, "MRC1"))
            Mrc16m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "MRC1")
        End If
        'GENSAN
        If GensanConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 35) = Val(GetFieldValueBranch(Qry, "GENSAN"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 36) = Val(GetFieldValueBranch(Qry, "GENSAN"))
            flxSP.TextMatrix(r, 37) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "GENSAN")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 38) = Val(GetFieldValueBranch(Qry, "GENSAN"))
            GenSan6M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "GENSAN")
        End If
        
        
        'MRC 2 GANTE
        If MRC2Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 39) = Val(GetFieldValueBranch(Qry, "GANTE"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 40) = Val(GetFieldValueBranch(Qry, "GANTE"))
            flxSP.TextMatrix(r, 41) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "GANTE")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 42) = Val(GetFieldValueBranch(Qry, "GANTE"))
            Mrc26M = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "GANTE")
        End If
        
        If YmarbelConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 43) = Val(GetFieldValueBranch(Qry, "MARBEL"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 44) = Val(GetFieldValueBranch(Qry, "MARBEL"))
            flxSP.TextMatrix(r, 45) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "MARBEL")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 46) = Val(GetFieldValueBranch(Qry, "MARBEL"))
            Ymarbel6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "MARBEL")
        End If
        
        If StotomasConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 47) = Val(GetFieldValueBranch(Qry, "STOTOMAS"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 48) = Val(GetFieldValueBranch(Qry, "STOTOMAS"))
            flxSP.TextMatrix(r, 49) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "STOTOMAS")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 50) = Val(GetFieldValueBranch(Qry, "STOTOMAS"))
            Stotomas6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "STOTOMAS")
        End If
        
        If SamalConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 51) = Val(GetFieldValueBranch(Qry, "SAMAL"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 52) = Val(GetFieldValueBranch(Qry, "SAMAL"))
            flxSP.TextMatrix(r, 53) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "SAMAL")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 54) = Val(GetFieldValueBranch(Qry, "SAMAL"))
            Samal6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "SAMAL")
        End If
        
        If Van3Connect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 55) = Val(GetFieldValueBranch(Qry, "VAN3"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 56) = Val(GetFieldValueBranch(Qry, "VAN3"))
            flxSP.TextMatrix(r, 57) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "VAN3")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 58) = Val(GetFieldValueBranch(Qry, "VAN3"))
            Van36m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "VAN3")
        End If
        
        If WsaleConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 59) = Val(GetFieldValueBranch(Qry, "WSALE"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 60) = Val(GetFieldValueBranch(Qry, "WSALE"))
            flxSP.TextMatrix(r, 61) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "WSALE")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 62) = Val(GetFieldValueBranch(Qry, "WSALE"))
            Wsale6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "WSALE")
        End If
        
        If CbntianConnect Then
            Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 59) = Val(GetFieldValueBranch(Qry, "WSALE"))
            Qry = ""
            Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 79) & ";"
            flxSP.TextMatrix(r, 60) = Val(GetFieldValueBranch(Qry, "WSALE"))
            flxSP.TextMatrix(r, 61) = GetSalesQtyBranch(Val(flxSP.TextMatrix(r, 79)), "WSALE")
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 79)) & ";"          ' AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            flxSP.TextMatrix(r, 62) = Val(GetFieldValueBranch(Qry, "WSALE"))
            Wsale6m = GetSalesQty6MBranch(Val(flxSP.TextMatrix(r, 79)), "WSALE")
        End If
        
        'Total Branch
        'SQ
        flxSP.TextMatrix(r, 67) = Val(flxSP.TextMatrix(r, 3)) + Val(flxSP.TextMatrix(r, 7)) + Val(flxSP.TextMatrix(r, 11)) + Val(flxSP.TextMatrix(r, 15)) + Val(flxSP.TextMatrix(r, 19)) + Val(flxSP.TextMatrix(r, 23)) + Val(flxSP.TextMatrix(r, 27)) + Val(flxSP.TextMatrix(r, 31)) + Val(flxSP.TextMatrix(r, 35)) + Val(flxSP.TextMatrix(r, 39)) + Val(flxSP.TextMatrix(r, 43)) + Val(flxSP.TextMatrix(r, 47)) + Val(flxSP.TextMatrix(r, 51)) + Val(flxSP.TextMatrix(r, 55)) + Val(flxSP.TextMatrix(r, 59))
        'BO
        flxSP.TextMatrix(r, 68) = Val(flxSP.TextMatrix(r, 4)) + Val(flxSP.TextMatrix(r, 8)) + Val(flxSP.TextMatrix(r, 12)) + Val(flxSP.TextMatrix(r, 16)) + Val(flxSP.TextMatrix(r, 20)) + Val(flxSP.TextMatrix(r, 24)) + Val(flxSP.TextMatrix(r, 28)) + Val(flxSP.TextMatrix(r, 32)) + Val(flxSP.TextMatrix(r, 36)) + Val(flxSP.TextMatrix(r, 40)) + Val(flxSP.TextMatrix(r, 44)) + Val(flxSP.TextMatrix(r, 48)) + Val(flxSP.TextMatrix(r, 52)) + Val(flxSP.TextMatrix(r, 56)) + Val(flxSP.TextMatrix(r, 60))
        'SQ
        flxSP.TextMatrix(r, 69) = Val(flxSP.TextMatrix(r, 5)) + Val(flxSP.TextMatrix(r, 9)) + Val(flxSP.TextMatrix(r, 13)) + Val(flxSP.TextMatrix(r, 17)) + Val(flxSP.TextMatrix(r, 21)) + Val(flxSP.TextMatrix(r, 25)) + Val(flxSP.TextMatrix(r, 29)) + Val(flxSP.TextMatrix(r, 33)) + Val(flxSP.TextMatrix(r, 37)) + Val(flxSP.TextMatrix(r, 41)) + Val(flxSP.TextMatrix(r, 45)) + Val(flxSP.TextMatrix(r, 49)) + Val(flxSP.TextMatrix(r, 53)) + Val(flxSP.TextMatrix(r, 57)) + Val(flxSP.TextMatrix(r, 61))
        'ST
        flxSP.TextMatrix(r, 70) = Val(flxSP.TextMatrix(r, 6)) + Val(flxSP.TextMatrix(r, 10)) + Val(flxSP.TextMatrix(r, 14)) + Val(flxSP.TextMatrix(r, 18)) + Val(flxSP.TextMatrix(r, 22)) + Val(flxSP.TextMatrix(r, 26)) + Val(flxSP.TextMatrix(r, 30)) + Val(flxSP.TextMatrix(r, 34)) + Val(flxSP.TextMatrix(r, 38)) + Val(flxSP.TextMatrix(r, 42)) + Val(flxSP.TextMatrix(r, 46)) + Val(flxSP.TextMatrix(r, 50)) + Val(flxSP.TextMatrix(r, 54)) + Val(flxSP.TextMatrix(r, 58)) + Val(flxSP.TextMatrix(r, 62))
        
        flxSP.TextMatrix(r, 74) = Ulas6M + Yamaha6M + Toril6M + Van16M + Van26M + Panabo6M + BinosPol6M + Mrc16m + GenSan6M + Mrc26M + Ymarbel6m + Stotomas6m + Samal6m + Van36m + Wsale6m + Cbntian6m + Val(flxSP.TextMatrix(r, 69))
        
        
        'suggested order Qty.  = 6m Sales Gante+6m Sales Qty. Branch -( Branch QH + Main QH) -  Back Order Main only
        'Before Jan. 4, 2013
        'spSOQ = Val(flxSP.TextMatrix(r, 42)) - (Val(flxSP.TextMatrix(r, 35)) + Val(flxSP.TextMatrix(r, 39))) - Val(flxSP.TextMatrix(r, 40))
        'new Monthly Avg (6Mos) * Order Mos  - (QH+BO)
        spSOQ = ((Val(flxSP.TextMatrix(r, 74)) / 6) * Val(cboMos.Text)) - (Val(flxSP.TextMatrix(r, 67)) + Val(flxSP.TextMatrix(r, 71)) + Val(flxSP.TextMatrix(r, 72)))
        flxSP.TextMatrix(r, 75) = IIf(spSOQ < 0, 0, Int(spSOQ + 0.49))
    Next r
    pBar.Visible = False
    'Call FillColorGrd
    'flxSP.Row = 0
    'flxSP.Col = 12
    flxSP.Col = 71: flxSP.Row = 0: flxSP.ColSel = 73: flxSP.RowSel = flxSP.Rows - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(150, 230, 150)
    If flxSP.Rows >= 1 Then
        'If flxSP.TextMatrix(0, 6) <> "" Then
        If flxSP.TextMatrix(0, 67) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
        
   cmdReplenish.Enabled = True
   cmdInsert.Enabled = True
   cmdOK.Enabled = True
   cmdCancel.Enabled = True
   'Timer1.Enabled = True
   Call fScrol_Change
   
End Sub

Private Sub cmdReplenish2_Click()
    Dim rsRpnsh As New ADODB.Recordset
    Dim Qry As String, r As Long, spSOQ As Double, rM As Integer
    Dim myW As String
    
    cmdReplenish2.Enabled = False
    cmdInsert.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    
    'identify conditions from Supplier and Classification
    If Trim(cboClass.Text) <> "" Then
        myW = "SPAREPARTS.REMARKS = \'" & cboClass.Text & "\'"
    Else
        myW = "SPAREPARTS.REMARKS is NOT NULL"
    End If
    If txtSuppCode.Text <> "" Then
        myW = myW & " AND SUPPLIERCODE = \'" & txtSuppCode.Text & "\'"
    End If
    
    DoEvents
    tabProcessing.Visible = True
    Call ResizeCompleteHead
    Call CollAlign
    Call DisableButtonOnProcess(True)
    Qry = "CALL mrcwarehouse.sp_podata('" & myW & "', '" & Format(SDfrom.Value, "yyyy-mm-dd") & "', '" & Format(SDto.Value, "yyyy-mm-dd") & "');"
    Set rsSP = New ADODB.Recordset
    If rsSP.State = adStateOpen Then rsSP.Close
    With rsSP
        DoEvents
        .Open Qry, LYRcon, adOpenStatic
        If .RecordCount = 1 Then
            .MoveFirst
            myW2 = .Fields(0)
            myW3 = .Fields(1)
            If myW2 = "Error" Then
                MsgBox myW3, vbCritical + vbOKOnly, "DB ERROR"
                .Close
                flxSP.Clear
                flxSP.Rows = 1
                Exit Sub
            End If
        End If
        If .RecordCount > 0 Then
            flxSP.Rows = 1
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount + 1
            .MoveFirst
            For rM = 1 To .RecordCount
                flxSP.Rows = rM
                flxSP.TextMatrix(rM - 1, 1) = !PartNumber
                flxSP.TextMatrix(rM - 1, 2) = !Description
                flxSP.TextMatrix(rM - 1, 3) = !Ulas_QH
                flxSP.TextMatrix(rM - 1, 4) = !Ulas_BO
                flxSP.TextMatrix(rM - 1, 5) = !Ulas_SQ
                flxSP.TextMatrix(rM - 1, 6) = !Ulas_ST
                flxSP.TextMatrix(rM - 1, 7) = !Rizal_QH
                flxSP.TextMatrix(rM - 1, 8) = !Rizal_BO
                flxSP.TextMatrix(rM - 1, 9) = !Rizal_SQ
                flxSP.TextMatrix(rM - 1, 10) = !Rizal_ST
                flxSP.TextMatrix(rM - 1, 11) = !Toril_QH
                flxSP.TextMatrix(rM - 1, 12) = !Toril_BO
                flxSP.TextMatrix(rM - 1, 13) = !Toril_SQ
                flxSP.TextMatrix(rM - 1, 14) = !Toril_ST
                flxSP.TextMatrix(rM - 1, 15) = !Van2_QH
                flxSP.TextMatrix(rM - 1, 16) = !Van2_BO
                flxSP.TextMatrix(rM - 1, 17) = !Van2_SQ
                flxSP.TextMatrix(rM - 1, 18) = !Van2_ST
                flxSP.TextMatrix(rM - 1, 19) = !Van56_QH
                flxSP.TextMatrix(rM - 1, 20) = !Van56_BO
                flxSP.TextMatrix(rM - 1, 21) = !Van56_SQ
                flxSP.TextMatrix(rM - 1, 22) = !Van56_ST
                flxSP.TextMatrix(rM - 1, 23) = !Panabo_QH
                flxSP.TextMatrix(rM - 1, 24) = !Panabo_BO
                flxSP.TextMatrix(rM - 1, 25) = !Panabo_SQ
                flxSP.TextMatrix(rM - 1, 26) = !Panabo_ST
                flxSP.TextMatrix(rM - 1, 31) = !Sob_QH
                flxSP.TextMatrix(rM - 1, 32) = !Sob_BO
                flxSP.TextMatrix(rM - 1, 33) = !Sob_SQ
                flxSP.TextMatrix(rM - 1, 34) = !Sob_ST
                flxSP.TextMatrix(rM - 1, 35) = !Gensan_QH
                flxSP.TextMatrix(rM - 1, 36) = !Gensan_BO
                flxSP.TextMatrix(rM - 1, 37) = !Gensan_SQ
                flxSP.TextMatrix(rM - 1, 38) = !Gensan_ST
                flxSP.TextMatrix(rM - 1, 39) = !Gante_QH
                flxSP.TextMatrix(rM - 1, 40) = !Gante_BO
                flxSP.TextMatrix(rM - 1, 41) = !Gante_SQ
                flxSP.TextMatrix(rM - 1, 42) = !Gante_ST
                flxSP.TextMatrix(rM - 1, 55) = !Van55_QH
                flxSP.TextMatrix(rM - 1, 56) = !Van55_BO
                flxSP.TextMatrix(rM - 1, 57) = !Van55_SQ
                flxSP.TextMatrix(rM - 1, 58) = !Van55_ST
                flxSP.TextMatrix(rM - 1, 59) = !WS_QH
                flxSP.TextMatrix(rM - 1, 60) = !WS_BO
                flxSP.TextMatrix(rM - 1, 61) = !WS_SQ
                flxSP.TextMatrix(rM - 1, 62) = !WS_ST
                flxSP.TextMatrix(rM - 1, 63) = !CAB_QH
                flxSP.TextMatrix(rM - 1, 64) = !CAB_BO
                flxSP.TextMatrix(rM - 1, 65) = !CAB_SQ
                flxSP.TextMatrix(rM - 1, 66) = !CAB_ST
                flxSP.TextMatrix(rM - 1, 71) = !WARE_QH
                flxSP.TextMatrix(rM - 1, 72) = !WARE_BO
                flxSP.TextMatrix(rM - 1, 73) = !WARE_ST 'NEW: IN-TRANSIT    old:6MSQ
                flxSP.TextMatrix(rM - 1, 77) = Format(!cost, "#,##0.00")
                flxSP.TextMatrix(rM - 1, 78) = IIf(IsNull(!ApplicModel), "", !ApplicModel)
                flxSP.TextMatrix(rM - 1, 79) = !SPAREPARTID
                flxSP.TextMatrix(rM - 1, 80) = IIf(IsNull(!DESCRIP2), "", !DESCRIP2)
                DoEvents
                pBar.Value = pBar.Value + 1
                .MoveNext
            Next rM
            
            flxSP.Visible = True
            If ShowOpt Then
                'flxSP.Top = 1920
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                'flxSP.Top = 1080
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
           ' Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
        Else
            flxSP.Clear
            flxSP.Rows = 1
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    Call AddInCombo
    pBar.Value = 0
    pBar.Max = flxSP.Rows
    For r = 0 To flxSP.Rows - 1
        pBar.Value = pBar.Value + 1
        Ulas6M = 0: Yamaha6M = 0: Toril6M = 0: Van16M = 0: Van26M = 0
        Mrc16m = 0: Panabo6M = 0: BinosPol6M = 0: GenSan6M = 0: Mrc26M = 0: Ymarbel6m = 0
        Stotomas6m = 0: Samal6m = 0: Van36m = 0: Wsale6m = 0: Cbntian6m = 0
        'Total Branch
        'SQ
        flxSP.TextMatrix(r, 67) = Val(flxSP.TextMatrix(r, 3)) + Val(flxSP.TextMatrix(r, 7)) + Val(flxSP.TextMatrix(r, 11)) + Val(flxSP.TextMatrix(r, 15)) + Val(flxSP.TextMatrix(r, 19)) + Val(flxSP.TextMatrix(r, 23)) + Val(flxSP.TextMatrix(r, 27)) + Val(flxSP.TextMatrix(r, 31)) + Val(flxSP.TextMatrix(r, 35)) + Val(flxSP.TextMatrix(r, 39)) + Val(flxSP.TextMatrix(r, 43)) + Val(flxSP.TextMatrix(r, 47)) + Val(flxSP.TextMatrix(r, 51)) + Val(flxSP.TextMatrix(r, 55)) + Val(flxSP.TextMatrix(r, 59)) + Val(flxSP.TextMatrix(r, 63))
        'BO
        flxSP.TextMatrix(r, 68) = Val(flxSP.TextMatrix(r, 4)) + Val(flxSP.TextMatrix(r, 8)) + Val(flxSP.TextMatrix(r, 12)) + Val(flxSP.TextMatrix(r, 16)) + Val(flxSP.TextMatrix(r, 20)) + Val(flxSP.TextMatrix(r, 24)) + Val(flxSP.TextMatrix(r, 28)) + Val(flxSP.TextMatrix(r, 32)) + Val(flxSP.TextMatrix(r, 36)) + Val(flxSP.TextMatrix(r, 40)) + Val(flxSP.TextMatrix(r, 44)) + Val(flxSP.TextMatrix(r, 48)) + Val(flxSP.TextMatrix(r, 52)) + Val(flxSP.TextMatrix(r, 56)) + Val(flxSP.TextMatrix(r, 60)) + Val(flxSP.TextMatrix(r, 64))
        'SQ
        flxSP.TextMatrix(r, 69) = Val(flxSP.TextMatrix(r, 5)) + Val(flxSP.TextMatrix(r, 9)) + Val(flxSP.TextMatrix(r, 13)) + Val(flxSP.TextMatrix(r, 17)) + Val(flxSP.TextMatrix(r, 21)) + Val(flxSP.TextMatrix(r, 25)) + Val(flxSP.TextMatrix(r, 29)) + Val(flxSP.TextMatrix(r, 33)) + Val(flxSP.TextMatrix(r, 37)) + Val(flxSP.TextMatrix(r, 41)) + Val(flxSP.TextMatrix(r, 45)) + Val(flxSP.TextMatrix(r, 49)) + Val(flxSP.TextMatrix(r, 53)) + Val(flxSP.TextMatrix(r, 57)) + Val(flxSP.TextMatrix(r, 61)) + Val(flxSP.TextMatrix(r, 65))
        'ST0ck transfer
        flxSP.TextMatrix(r, 70) = Val(flxSP.TextMatrix(r, 6)) + Val(flxSP.TextMatrix(r, 10)) + Val(flxSP.TextMatrix(r, 14)) + Val(flxSP.TextMatrix(r, 18)) + Val(flxSP.TextMatrix(r, 22)) + Val(flxSP.TextMatrix(r, 26)) + Val(flxSP.TextMatrix(r, 30)) + Val(flxSP.TextMatrix(r, 34)) + Val(flxSP.TextMatrix(r, 38)) + Val(flxSP.TextMatrix(r, 42)) + Val(flxSP.TextMatrix(r, 46)) + Val(flxSP.TextMatrix(r, 50)) + Val(flxSP.TextMatrix(r, 54)) + Val(flxSP.TextMatrix(r, 58)) + Val(flxSP.TextMatrix(r, 62)) + Val(flxSP.TextMatrix(r, 66))
        'Total Qty OnHand (TTLQH)
        flxSP.TextMatrix(r, 74) = Val(flxSP.TextMatrix(r, 67)) + Val(flxSP.TextMatrix(r, 71))
                'Branch Total Sales Qty(69) Less TtlQH(74)+Warehouse InTransit(73) + Warehouse BO(72)
        spSOQ = Val(flxSP.TextMatrix(r, 69)) - (Val(flxSP.TextMatrix(r, 74)) + Val(flxSP.TextMatrix(r, 73)) + Val(flxSP.TextMatrix(r, 72)))
        flxSP.TextMatrix(r, 75) = IIf(spSOQ < 0, 0, Int(spSOQ + 0.49))

'        flxSP.TextMatrix(r, 74) = Ulas6M + Yamaha6M + Toril6M + Van16M + Van26M + Panabo6M + BinosPol6M + Mrc16m + GenSan6M + Mrc26M + Ymarbel6m + Stotomas6m + Samal6m + Van36m + Wsale6m + Cbntian6m + Val(flxSP.TextMatrix(r, 73))
'
'        'suggested order Qty.  = 6m Sales Gante+6m Sales Qty. Branch -( Branch QH + Main QH) -  Back Order Main only
'        'Before Jan. 4, 2013
'        'spSOQ = Val(flxSP.TextMatrix(r, 42)) - (Val(flxSP.TextMatrix(r, 35)) + Val(flxSP.TextMatrix(r, 39))) - Val(flxSP.TextMatrix(r, 40))
'        'new Monthly Avg (6Mos) * Order Mos  - (QH+BO)
'        spSOQ = ((Val(flxSP.TextMatrix(r, 74) / 6)) * Val(cbomos.Text)) - (Val(flxSP.TextMatrix(r, 67)) + Val(flxSP.TextMatrix(r, 71)) + Val(flxSP.TextMatrix(r, 72)))
'        flxSP.TextMatrix(r, 75) = IIf(spSOQ < 0, 0, Int(spSOQ + 0.49))
        
    Next r
    pBar.Visible = False
    flxSP.Col = 71: flxSP.Row = 0: flxSP.ColSel = 73: flxSP.RowSel = flxSP.Rows - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = RGB(150, 230, 150)
    If flxSP.Rows >= 1 Then
        'If flxSP.TextMatrix(1, 3) <> "" Then
                         'QTY on Hand
        If flxSP.TextMatrix(0, 67) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
    tabProcessing.Visible = False
    Call DisableButtonOnProcess(False)


End Sub

Private Sub flxSP_Click()
    
    ActRow = flxSP.Row
    ActCol = flxSP.Col
    If ActCol <> 76 Then: RwSel = False: Exit Sub
    With txtOrder
        .Top = flxSP.CellTop + flxSP.Top
        .Left = flxSP.CellLeft + flxSP.Left
        .Width = flxSP.CellWidth
        .Text = Trim(flxSP.Text)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With
    RwSel = True
    flxSP.RowSel = flxSP.Row
    flxSP.Col = 0
    flxSP.ColSel = 80
    RwSel = False
    flxSP.FillStyle = flexFillRepeat
    
End Sub
Private Sub flxSP_DblClick()
    If flxSP.TextMatrix(flxSP.Row, 1) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
        Call mnuMark_Click
    Else
        Call mnuUnmark_Click
    End If
End Sub

Private Sub flxSP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxSP.TextMatrix(flxSP.Row, 1) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
    ElseIf KeyCode = vbKeyF11 Then
        flxSP.Col = 76
        Call flxSP_Click
    End If
End Sub
Private Sub flxSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSP_Click
End Sub
Private Sub flxSP_LeaveCell()
    If RwSel = True Then Exit Sub
    flxSP.ColSel = 80
    flxSP.FillStyle = flexFillRepeat
End Sub
Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Val(flxSP.TextMatrix(flxSP.Row, 1)) = 0 Then Exit Sub
    If Button = vbRightButton Then
        If flxSP.TextMatrix(flxSP.Row, 80) = "INACTIVE" Then
            mnuInAct.Visible = False
            mnuAct.Visible = True
        Else
            mnuInAct.Visible = True
            mnuAct.Visible = False
        End If
        Call PopupMenu(mnuSP)
    End If
End Sub

Private Sub flxSP_Scroll()
    txtOrder.Visible = False
    txtOrder.Text = ""
End Sub

Private Sub flxSP_SelChange()

    If flxSP.Col = 76 Then Exit Sub
        RwSel = True
        flxSP.RowSel = flxSP.Row
        flxSP.Col = 0
        flxSP.ColSel = 80
        RwSel = False
        flxSP.FillStyle = flexFillRepeat
    
End Sub

Private Sub flxSPH_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnurg)
    End If
End Sub


Private Sub Form_Load()
    Dim Qry As String, m As Integer
    
    RwSel = True
    flxSP.Cols = 82
    Call cmdOptHide_Click
    fScrol.Left = flxSPH.ColPos(6) + 80
    With SPordering
        If .LockFlx = "True" Then
            cmdInsert.Enabled = False
            cmdOK.Enabled = False
        End If
        'SDfrom.Value = .SalesDateFrom.Value
        SDfrom.Value = DateAdd("m", -6, Date)
        SDto.Value = .SalesDateTo.Value
        cboMos.Text = 6
        For m = 1 To 12
            cboMos.AddItem m
        Next m
        'cboMos.AddItem "All"
    End With
    
    'Set Search Description
    updRef.Value = 1
    
    Call ResizeCompleteHead
    Call ResizeCompleteFlxSP
    
    'Connect to Branches
'    If ConnectUlas Then
'        UlasConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon.Caption = "ULAS: " & GetFieldValueBranch(Qry, "ULAS")
'    Else
'        UlasConnect = False
'        lblBrnCon.Caption = "ULAS: Disconnected"
'    End If
'    If ConnectYamaha Then
'        YamahaConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflowc.CFdate,'%c/%d/%Y') From cashflowc Order By cashflowc.CFdate Desc;"
'        lblBrnCon.Caption = lblBrnCon.Caption & ", YAMAHA: " & GetFieldValueBranch(Qry, "YAMAHA")
'    Else
'        YamahaConnect = False
'        lblBrnCon.Caption = lblBrnCon.Caption & ", YAMAHA: Disconnected"
'    End If
'    If ConnectToril Then
'        TorilConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon.Caption = lblBrnCon.Caption & ", DEPOT: " & GetFieldValueBranch(Qry, "TORIL")
'    Else
'        TorilConnect = False
'        lblBrnCon.Caption = lblBrnCon.Caption & ", DEPOT: Disconnected"
'    End If
'    If ConnectVAN1 Then
'        Van1Connect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon.Caption = lblBrnCon.Caption & ", VAN1: " & GetFieldValueBranch(Qry, "VAN1")
'    Else
'        Van1Connect = False
'        lblBrnCon.Caption = lblBrnCon.Caption & ", VAN1: Disconnected"
'    End If
'    If ConnectVAN2 Then
'        Van2Connect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon.Caption = lblBrnCon.Caption & ", VAN2: " & GetFieldValueBranch(Qry, "VAN2")
'    Else
'        Van2Connect = False
'        lblBrnCon.Caption = lblBrnCon.Caption & ", VAN2: Disconnected"
'    End If
'    If ConnectPanabo Then
'        PanaboConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon.Caption = lblBrnCon.Caption & ", PANABO: " & GetFieldValueBranch(Qry, "PANABO")
'    Else
'        PanaboConnect = False
'        lblBrnCon.Caption = lblBrnCon.Caption & ", PANABO: Disconnected"
'    End If
'    If ConnectHondaPol Then
'        BinosPolConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflowc.CFdate,'%c/%d/%Y') From cashflowc Order By cashflowc.CFdate Desc;"
'        lblBrnCon.Caption = lblBrnCon.Caption & ", POL: " & GetFieldValueBranch(Qry, "POLOMOLOK")
'    Else
'        BinosPolConnect = False
'        lblBrnCon.Caption = lblBrnCon.Caption & ", POL: Disconnected"
'    End If
'    If ConnectMRC1 Then
'        MRC1Connect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon.Caption = lblBrnCon.Caption & ", MRC1: " & GetFieldValueBranch(Qry, "MRC1")
'    Else
'        MRC1Connect = False
'        lblBrnCon.Caption = lblBrnCon.Caption & ", MRC1: Disconnected"
'    End If
'    If ConnectMRCGensan Then
'        GensanConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon.Caption = lblBrnCon.Caption & ", GENSAN: " & GetFieldValueBranch(Qry, "GENSAN")
'    Else
'        GensanConnect = False
'        lblBrnCon.Caption = lblBrnCon.Caption & ", GENSAN: Disconnected"
'    End If
'    If ConnectYmarbel Then
'        YmarbelConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon2.Caption = "MARBEL: " & GetFieldValueBranch(Qry, "MARBEL")
'    Else
'        YmarbelConnect = False
'        lblBrnCon2.Caption = "MARBEL: Disconnected"
'    End If
'    If ConnectMRC2 Then
'        MRC2Connect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", GANTE: " & GetFieldValueBranch(Qry, "GANTE")
'    Else
'        MRC2Connect = False
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", GANTE: Disconnected"
'    End If
'    If ConnectStotomas Then
'        StotomasConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", STOTOMAS: " & GetFieldValueBranch(Qry, "STOTOMAS")
'    Else
'        StotomasConnect = False
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", STOTOMAS: Disconnected"
'    End If
'    If ConnectSamal Then
'        SamalConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", SAMAL: " & GetFieldValueBranch(Qry, "SAMAL")
'    Else
'        SamalConnect = False
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", SAMAL: Disconnected"
'    End If
'    If ConnectVan3 Then
'        Van3Connect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", VAN3: " & GetFieldValueBranch(Qry, "VAN3")
'    Else
'        Van3Connect = False
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", VAN3: Disconnected"
'    End If
'    If ConnectWsale Then
'        WsaleConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", WSALE: " & GetFieldValueBranch(Qry, "WSALE")
'    Else
'        WsaleConnect = False
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", WSALE: Disconnected"
'    End If
'    If ConnectCabantian Then
'        CbntianConnect = True
'        Qry = "SELECT DATE_FORMAT(cashflow.CFdate,'%c/%d/%Y') From cashflow Order By cashflow.CFdate Desc;"
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", CABANTIAN: " & GetFieldValueBranch(Qry, "CABANTIAN")
'    Else
'        CbntianConnect = False
'        lblBrnCon2.Caption = lblBrnCon2.Caption & ", CABANTIAN: Disconnected"
'    End If
    
    'Call ConnectMU
    'Call ConnectYM
    'Call ConnectMD
    'Call ConnectVAN
    'Call ConnectMRC
    
    fScrol.Top = flxSPH.Top + 15
    fScrol.Height = flxSPH.RowHeight(0)
    fScrol.Left = flxSPH.ColPos(3) + 80
    fScrol.ZOrder
    
    
'    fScrol.Top = flxSPH.Top + 15
'    fScrol.Height = flxSPH.RowHeight(0)
'    fScrol.Left = flxSPH.ColPos(6) + 80
    
    Timer1.Enabled = True
    'LYRmain.sBar.Panels(6).Text = "Find Spare Parts"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'LYRmain.sBar.Panels(6).Text = ""
    'MainCon.Close
    'BrchConMU.Close
    'BrchConYM.Close
    'BrchConMD.Close
    'BrchConVAN.Close
    'BrnConMRC2.Close
    'Timer1.Enabled = False
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub

Private Sub fScrol_Change()
    Dim reAlns As Integer, aCol As Integer
    
    With flxSP
       If fScrol.Value = 1 Then 'moto-r ulas
            .ColWidth(3) = (.Width - 300) * 0.036
            .ColWidth(4) = (.Width - 300) * 0.036
            .ColWidth(5) = (.Width - 300) * 0.036
            .ColWidth(6) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(3) = (.Width - 300) * 0.036
            flxSPH.ColWidth(4) = (.Width - 300) * 0.036
            flxSPH.ColWidth(5) = (.Width - 300) * 0.036
            flxSPH.ColWidth(6) = (.Width - 300) * 0.036
            For reCol = 3 To 6
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            
            For reCol = 7 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
        ElseIf fScrol.Value = 2 Then    'yamaha
            .ColWidth(7) = (.Width - 300) * 0.036
            .ColWidth(8) = (.Width - 300) * 0.036
            .ColWidth(9) = (.Width - 300) * 0.036
            .ColWidth(10) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(7) = (.Width - 300) * 0.036
            flxSPH.ColWidth(8) = (.Width - 300) * 0.036
            flxSPH.ColWidth(9) = (.Width - 300) * 0.036
            flxSPH.ColWidth(10) = (.Width - 300) * 0.036
            For reCol = 7 To 10
                 flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 6
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
            For reCol = 11 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
        ElseIf fScrol.Value = 3 Then    'MCParts Depot
            .ColWidth(11) = (.Width - 300) * 0.036
            .ColWidth(12) = (.Width - 300) * 0.036
            .ColWidth(13) = (.Width - 300) * 0.036
            .ColWidth(14) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(11) = (.Width - 300) * 0.036
            flxSPH.ColWidth(12) = (.Width - 300) * 0.036
            flxSPH.ColWidth(13) = (.Width - 300) * 0.036
            flxSPH.ColWidth(14) = (.Width - 300) * 0.036
            For reCol = 11 To 14
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 10
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 15 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        
        ElseIf fScrol.Value = 4 Then    'van selling I
            .ColWidth(15) = (.Width - 300) * 0.036
            .ColWidth(16) = (.Width - 300) * 0.036
            .ColWidth(17) = (.Width - 300) * 0.036
            .ColWidth(18) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(15) = (.Width - 300) * 0.036
            flxSPH.ColWidth(16) = (.Width - 300) * 0.036
            flxSPH.ColWidth(17) = (.Width - 300) * 0.036
            flxSPH.ColWidth(18) = (.Width - 300) * 0.036
            For reCol = 15 To 18
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 14
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 19 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        ElseIf fScrol.Value = 5 Then    'Van Selling II
            .ColWidth(19) = (.Width - 300) * 0.036
            .ColWidth(20) = (.Width - 300) * 0.036
            .ColWidth(21) = (.Width - 300) * 0.036
            .ColWidth(22) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(19) = (.Width - 300) * 0.036
            flxSPH.ColWidth(20) = (.Width - 300) * 0.036
            flxSPH.ColWidth(21) = (.Width - 300) * 0.036
            flxSPH.ColWidth(22) = (.Width - 300) * 0.036
            For reCol = 19 To 22
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 18
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 23 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        ElseIf fScrol.Value = 6 Then    'MRC Panabo
            .ColWidth(23) = (.Width - 300) * 0.036
            .ColWidth(24) = (.Width - 300) * 0.036
            .ColWidth(25) = (.Width - 300) * 0.036
            .ColWidth(26) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(23) = (.Width - 300) * 0.036
            flxSPH.ColWidth(24) = (.Width - 300) * 0.036
            flxSPH.ColWidth(25) = (.Width - 300) * 0.036
            flxSPH.ColWidth(26) = (.Width - 300) * 0.036
            For reCol = 23 To 26
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 22
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 27 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        ElseIf fScrol.Value = 7 Then    'MRC POLOMOLOK
            .ColWidth(27) = (.Width - 300) * 0.036
            .ColWidth(28) = (.Width - 300) * 0.036
            .ColWidth(29) = (.Width - 300) * 0.036
            .ColWidth(30) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(27) = (.Width - 300) * 0.036
            flxSPH.ColWidth(28) = (.Width - 300) * 0.036
            flxSPH.ColWidth(29) = (.Width - 300) * 0.036
            flxSPH.ColWidth(30) = (.Width - 300) * 0.036
            For reCol = 27 To 30
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 26
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 31 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        ElseIf fScrol.Value = 8 Then    'MRC SOBRECARY
            .ColWidth(31) = (.Width - 300) * 0.036
            .ColWidth(32) = (.Width - 300) * 0.036
            .ColWidth(33) = (.Width - 300) * 0.036
            .ColWidth(34) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(31) = (.Width - 300) * 0.036
            flxSPH.ColWidth(32) = (.Width - 300) * 0.036
            flxSPH.ColWidth(33) = (.Width - 300) * 0.036
            flxSPH.ColWidth(34) = (.Width - 300) * 0.036
            For reCol = 31 To 34
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 30
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 35 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        
        ElseIf fScrol.Value = 9 Then    'MRC GENSAN
            .ColWidth(35) = (.Width - 300) * 0.036
            .ColWidth(36) = (.Width - 300) * 0.036
            .ColWidth(37) = (.Width - 300) * 0.036
            .ColWidth(38) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(35) = (.Width - 300) * 0.036
            flxSPH.ColWidth(36) = (.Width - 300) * 0.036
            flxSPH.ColWidth(37) = (.Width - 300) * 0.036
            flxSPH.ColWidth(38) = (.Width - 300) * 0.036
            For reCol = 35 To 38
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 34
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 39 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
        ElseIf fScrol.Value = 10 Then    'MRC 2 RETAIL
            .ColWidth(39) = (.Width - 300) * 0.036
            .ColWidth(40) = (.Width - 300) * 0.036
            .ColWidth(41) = (.Width - 300) * 0.036
            .ColWidth(42) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(39) = (.Width - 300) * 0.036
            flxSPH.ColWidth(40) = (.Width - 300) * 0.036
            flxSPH.ColWidth(41) = (.Width - 300) * 0.036
            flxSPH.ColWidth(42) = (.Width - 300) * 0.036
            
            For reCol = 39 To 42
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 38
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 43 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
        ElseIf fScrol.Value = 11 Then    'YAMAHA MARBEL
            .ColWidth(43) = (.Width - 300) * 0.036
            .ColWidth(44) = (.Width - 300) * 0.036
            .ColWidth(45) = (.Width - 300) * 0.036
            .ColWidth(46) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(43) = (.Width - 300) * 0.036
            flxSPH.ColWidth(44) = (.Width - 300) * 0.036
            flxSPH.ColWidth(45) = (.Width - 300) * 0.036
            flxSPH.ColWidth(46) = (.Width - 300) * 0.036
            
            For reCol = 43 To 46
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 42
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 47 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
        ElseIf fScrol.Value = 12 Then    'BINOS STO TOMAS
            .ColWidth(47) = (.Width - 300) * 0.036
            .ColWidth(48) = (.Width - 300) * 0.036
            .ColWidth(49) = (.Width - 300) * 0.036
            .ColWidth(50) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(47) = (.Width - 300) * 0.036
            flxSPH.ColWidth(48) = (.Width - 300) * 0.036
            flxSPH.ColWidth(49) = (.Width - 300) * 0.036
            flxSPH.ColWidth(50) = (.Width - 300) * 0.036
            
            For reCol = 47 To 50
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 46
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 51 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
        ElseIf fScrol.Value = 13 Then    'BINOS SAMAL
            .ColWidth(51) = (.Width - 300) * 0.036
            .ColWidth(52) = (.Width - 300) * 0.036
            .ColWidth(53) = (.Width - 300) * 0.036
            .ColWidth(54) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(51) = (.Width - 300) * 0.036
            flxSPH.ColWidth(52) = (.Width - 300) * 0.036
            flxSPH.ColWidth(53) = (.Width - 300) * 0.036
            flxSPH.ColWidth(54) = (.Width - 300) * 0.036
            
            For reCol = 51 To 54
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 50
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 55 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
        ElseIf fScrol.Value = 14 Then    'VAN3
            .ColWidth(55) = (.Width - 300) * 0.036
            .ColWidth(56) = (.Width - 300) * 0.036
            .ColWidth(57) = (.Width - 300) * 0.036
            .ColWidth(58) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(55) = (.Width - 300) * 0.036
            flxSPH.ColWidth(56) = (.Width - 300) * 0.036
            flxSPH.ColWidth(57) = (.Width - 300) * 0.036
            flxSPH.ColWidth(58) = (.Width - 300) * 0.036
            
            For reCol = 55 To 58
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 54
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 59 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        
        ElseIf fScrol.Value = 15 Then    'WHOLESALE
            .ColWidth(59) = (.Width - 300) * 0.036
            .ColWidth(60) = (.Width - 300) * 0.036
            .ColWidth(61) = (.Width - 300) * 0.036
            .ColWidth(62) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(59) = (.Width - 300) * 0.036
            flxSPH.ColWidth(60) = (.Width - 300) * 0.036
            flxSPH.ColWidth(61) = (.Width - 300) * 0.036
            flxSPH.ColWidth(62) = (.Width - 300) * 0.036
            
            For reCol = 59 To 62
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 58
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            For reCol = 63 To 66
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
        ElseIf fScrol.Value = 16 Then    'cabantian
            .ColWidth(63) = (.Width - 300) * 0.036
            .ColWidth(64) = (.Width - 300) * 0.036
            .ColWidth(65) = (.Width - 300) * 0.036
            .ColWidth(66) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(63) = (.Width - 300) * 0.036
            flxSPH.ColWidth(64) = (.Width - 300) * 0.036
            flxSPH.ColWidth(65) = (.Width - 300) * 0.036
            flxSPH.ColWidth(66) = (.Width - 300) * 0.036
            
            For reCol = 63 To 66
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 62
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        End If
    End With
        
End Sub
Private Sub fScrol_Scroll()
    'Timer1.Enabled = False
    Call fScrol_Change
    'Timer1.Enabled = True
End Sub

Private Sub mnuAct_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want ACTIVATE Partnumber " & flxSP.TextMatrix(flxSP.Row, 1) & "?", vbQuestion + vbYesNo, "In-Active")
    If Ans = vbNo Then Exit Sub
    
    flxSP.Col = 1:  flxSP.ColSel = 80: flxSP.RowSel = flxSP.Row: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(0, 0, 0)
    flxSP.TextMatrix(flxSP.Row, 80) = ""
    
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & ";"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 79)) & ";"

End Sub

Private Sub mnuInAct_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want Partnumber " & flxSP.TextMatrix(flxSP.Row, 1) & " to be IN-ACTIVE?", vbQuestion + vbYesNo, "In-Active")
    If Ans = vbNo Then Exit Sub
    
    flxSP.Col = 1:  flxSP.ColSel = 80: flxSP.RowSel = flxSP.Row: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(255, 0, 0)
    flxSP.TextMatrix(flxSP.Row, 80) = "INACTIVE"
    
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = 'INACTIVE' WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = 'INACTIVE' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 79)) & ";"
    
End Sub

Private Sub mnuMark_Click()
    With flxSP
        .Col = 0
        .CellFontName = "WINGDINGS"
        .CellFontSize = 12
        .Text = lblChk.Caption
        .CellAlignment = 4
    End With
End Sub

Private Sub mnuMarkAll_Click()
    Dim i As Integer
    
    With flxSP
        For i = 1 To .Rows - 1
            .Row = i
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .TextMatrix(i, 0) = lblChk.Caption
        Next i
    End With
    
End Sub

Private Sub mnuUnmark_Click()
    flxSP.TextMatrix(flxSP.Row, 0) = ""
End Sub

Private Sub mnuUnmarkAll_Click()
    Dim i As Integer
    
    For i = 1 To flxSP.Rows - 1
        flxSP.TextMatrix(i, 0) = ""
    Next i
End Sub


Private Sub OptrBO_Change()
    Select Case OptrBO.Value
        Case 0
            lblOprBO.Caption = "="
        Case 1
            lblOprBO.Caption = ">"
        Case 2
            lblOprBO.Caption = ">="
        Case 3
            lblOprBO.Caption = "<"
        Case 4
            lblOprBO.Caption = "<="
        Case 5
            lblOprBO.Caption = "<>"
    End Select

End Sub
Private Sub OptrQH_Change()
    Select Case OptrQH.Value
        Case 0
            lblOprQH.Caption = "="
        Case 1
            lblOprQH.Caption = ">"
        Case 2
            lblOprQH.Caption = ">="
        Case 3
            lblOprQH.Caption = "<"
        Case 4
            lblOprQH.Caption = "<="
        Case 5
            lblOprQH.Caption = "<>"
    End Select
End Sub
Private Sub OptrSQ_Change()
    Select Case OptrSQ.Value
        Case 0
            lblOprSQ.Caption = "="
        Case 1
            lblOprSQ.Caption = ">"
        Case 2
            lblOprSQ.Caption = ">="
        Case 3
            lblOprSQ.Caption = "<"
        Case 4
            lblOprSQ.Caption = "<="
        Case 5
            lblOprSQ.Caption = "<>"
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim Qry As String
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    
    cboSupplier.Text = SPordering.cboSupplier.Text
    txtSuppCode.Text = SPordering.txtScode.Text
   
   
    Qry = "SELECT DISTINCT spareparts.REMARKS From spareparts Order By spareparts.REMARKS"
    With rsClass
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboClass.Clear
            Do While Not .EOF
                cboClass.AddItem !Remarks
                .MoveNext
            Loop
        End If
    End With


'    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
'
'    For Cw = 0 To 24
'        flxSP.ColWidth(Cw) = flxSPH.ColWidth(Cw)
'    Next Cw
'    fScrol.Left = flxSPH.ColPos(6) + 60
'    flxSP.ColWidth(26) = flxSPH.ColWidth(25)
'    flxSP.ColWidth(27) = flxSPH.ColWidth(26)
'    tm = tm + 1
'    If tm = 60 Then chkResize.Value = 0
    Timer1.Enabled = False
End Sub
Private Sub txtOrder_GotFocus()
    'Timer1.Enabled = False
End Sub
Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 75)) And UserID <> 2 Then
            flxSP.TextMatrix(ActRow, 76) = ""
            flxSP.TextMatrix(ActRow, 0) = ""
            txtOrder.Visible = False
            flxSP.Col = 76
            Exit Sub
        ElseIf Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 75)) And UserID = 2 Then
            MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            txtOrder.SetFocus
            Exit Sub
        ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 75)) And Val(txtOrder.Text) <> 0 Then
            MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            txtOrder.SetFocus
            Exit Sub
        End If
        txtOrder.Text = flxSP.TextMatrix(ActRow, 76)
        flxSP.SetFocus
        flxSP.Row = ActRow
        flxSP.Col = 76
        txtOrder.Visible = False
    End If
End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxSP.SetFocus
        flxSP.Row = ActRow
        flxSP.Col = 76
   End If
End Sub
Private Sub txtOrder_LostFocus()
    With flxSP
    
        'If Val(txtOrder.Text) > Val(.TextMatrix(ActRow, 47)) And (UserID <> 4 Or UserID <> 5) Then
'        If Val(txtOrder.Text) > Val(.TextMatrix(ActRow, 47)) Then
'            MsgBox ("Can't add Order Quantity" & vbCrLf & _
'                    "which Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
'            flxSP.TextMatrix(ActRow, 48) = ""
'            flxSP.TextMatrix(ActRow, 0) = ""
'            If txtOrder.Visible = True Then txtOrder.SetFocus ': SendKeys "{Home}+{End}"
'            Exit Sub
'        End If  'Else           'isumpay kay temporary lang pra kay iva
        If Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 75)) Then
            MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            If UserID <> 2 Then
                .TextMatrix(ActRow, 76) = ""
                .TextMatrix(ActRow, 0) = ""
                .Col = 76
                Exit Sub    'Requested by sir Jo 6/10/2016
            End If
        ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 75)) And Val(txtOrder.Text) <> 0 Then
            MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        End If
        
        .TextMatrix(ActRow, 76) = txtOrder.Text
        txtOrder.Visible = False
        If Val(.TextMatrix(ActRow, 76)) > 0 Then
            .Row = ActRow
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .Col = 76
        Else
            .TextMatrix(ActRow, 76) = ""
            .TextMatrix(ActRow, 0) = ""
            .Col = 76
        End If
        
        'Timer1.Enabled = True
    End With
End Sub
Private Sub UpDRef_Change()
    If updRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    Else
        lblRefSearch.Caption = "FIND PARTS DESC."
    End If
   
End Sub
Private Sub ResizeCompleteHead()
    Dim aCol As Integer, aCel As Integer, reAln As Integer
    Dim i As Integer
    
    With flxSPH
        
        For aCol = 1 To 77
            .ColAlignment(aCol) = 4
        Next aCol
      '  For reAln = 7 To 14
      '      .ColAlignment(reAln) = 1
      '  Next
        
        .TextMatrix(0, 1) = "S P A R E P A R T S"
        .TextMatrix(0, 2) = "S P A R E P A R T S"
        
        .TextMatrix(0, 3) = "MOTO-R ULAS"
        .TextMatrix(0, 4) = "MOTO-R ULAS"
        .TextMatrix(0, 5) = "MOTO-R ULAS"
        .TextMatrix(0, 6) = "MOTO-R ULAS"
        
        .TextMatrix(0, 7) = "YM MOTORHUB"
        .TextMatrix(0, 8) = "YM MOTORHUB"
        .TextMatrix(0, 9) = "YM MOTORHUB"
        .TextMatrix(0, 10) = "YM MOTORHUB"
        
        .TextMatrix(0, 11) = "MCP DEPOT"
        .TextMatrix(0, 12) = "MCP DEPOT"
        .TextMatrix(0, 13) = "MCP DEPOT"
        .TextMatrix(0, 14) = "MCP DEPOT"
             
        .TextMatrix(0, 15) = "VAN I"
        .TextMatrix(0, 16) = "VAN I"
        .TextMatrix(0, 17) = "VAN I"
        .TextMatrix(0, 18) = "VAN I"
             
        .TextMatrix(0, 19) = "VAN II"
        .TextMatrix(0, 20) = "VAN II"
        .TextMatrix(0, 21) = "VAN II"
        .TextMatrix(0, 22) = "VAN II"
             
        .TextMatrix(0, 23) = "MRC PANABO"
        .TextMatrix(0, 24) = "MRC PANABO"
        .TextMatrix(0, 25) = "MRC PANABO"
        .TextMatrix(0, 26) = "MRC PANABO"
                  
        .TextMatrix(0, 27) = "POLOMOLOK"
        .TextMatrix(0, 28) = "POLOMOLOK"
        .TextMatrix(0, 29) = "POLOMOLOK"
        .TextMatrix(0, 30) = "POLOMOLOK"
                  
        .TextMatrix(0, 31) = "MRC SOBRECARY"
        .TextMatrix(0, 32) = "MRC SOBRECARY"
        .TextMatrix(0, 33) = "MRC SOBRECARY"
        .TextMatrix(0, 34) = "MRC SOBRECARY"
        
        .TextMatrix(0, 35) = "MRC GENSAN"
        .TextMatrix(0, 36) = "MRC GENSAN"
        .TextMatrix(0, 37) = "MRC GENSAN"
        .TextMatrix(0, 38) = "MRC GENSAN"
        
        .TextMatrix(0, 39) = "MRC 2 RETAIL"
        .TextMatrix(0, 40) = "MRC 2 RETAIL"
        .TextMatrix(0, 41) = "MRC 2 RETAIL"
        .TextMatrix(0, 42) = "MRC 2 RETAIL"
        
        .TextMatrix(0, 43) = "YAMAHA MARBEL"
        .TextMatrix(0, 44) = "YAMAHA MARBEL"
        .TextMatrix(0, 45) = "YAMAHA MARBEL"
        .TextMatrix(0, 46) = "YAMAHA MARBEL"
        
        .TextMatrix(0, 47) = "BINOS STO.TOMAS"
        .TextMatrix(0, 48) = "BINOS STO.TOMAS"
        .TextMatrix(0, 49) = "BINOS STO.TOMAS"
        .TextMatrix(0, 50) = "BINOS STO.TOMAS"
        
        .TextMatrix(0, 51) = "BINOS SAMAL"
        .TextMatrix(0, 52) = "BINOS SAMAL"
        .TextMatrix(0, 53) = "BINOS SAMAL"
        .TextMatrix(0, 54) = "BINOS SAMAL"
        
        .TextMatrix(0, 55) = "MRC VAN3"
        .TextMatrix(0, 56) = "MRC VAN3"
        .TextMatrix(0, 57) = "MRC VAN3"
        .TextMatrix(0, 58) = "MRC VAN3"
        
        .TextMatrix(0, 59) = "MRC WHOLESALE"
        .TextMatrix(0, 60) = "MRC WHOLESALE"
        .TextMatrix(0, 61) = "MRC WHOLESALE"
        .TextMatrix(0, 62) = "MRC WHOLESALE"
        
        .TextMatrix(0, 63) = "MRC CABANTIAN"
        .TextMatrix(0, 64) = "MRC CABANTIAN"
        .TextMatrix(0, 65) = "MRC CABANTIAN"
        .TextMatrix(0, 66) = "MRC CABANTIAN"
             
        .TextMatrix(0, 67) = "BRANCH TOTAL"
        .TextMatrix(0, 68) = "BRANCH TOTAL"
        .TextMatrix(0, 69) = "BRANCH TOTAL"
        .TextMatrix(0, 70) = "BRANCH TOTAL"
        
        .TextMatrix(0, 71) = "MRC WAREHOUSE"
        .TextMatrix(0, 72) = "MRC WAREHOUSE"
        .TextMatrix(0, 73) = "MRC WAREHOUSE"
        
        .TextMatrix(0, 74) = "PURCHASE ORDER"
        .TextMatrix(0, 75) = "PURCHASE ORDER"
        .TextMatrix(0, 76) = "PURCHASE ORDER"
        .TextMatrix(0, 77) = "PURCHASE ORDER"
        
        
        'moto-r Ulas
        .TextMatrix(1, 1) = "PARTNUMBER"
        .TextMatrix(1, 2) = "DESCRIPTION"
        .TextMatrix(1, 3) = "QH"
        .TextMatrix(1, 4) = "BO"
        .TextMatrix(1, 5) = "SQ"
        .TextMatrix(1, 6) = "ST"
        'Yamaha Motorhub
        .TextMatrix(1, 7) = "QH"
        .TextMatrix(1, 8) = "BO"
        .TextMatrix(1, 9) = "SQ"
        .TextMatrix(1, 10) = "ST"
        'MC Parts Depot
        .TextMatrix(1, 11) = "QH"
        .TextMatrix(1, 12) = "BO"
        .TextMatrix(1, 13) = "SQ"
        .TextMatrix(1, 14) = "ST"
        'VAN I
        .TextMatrix(1, 15) = "QH"
        .TextMatrix(1, 16) = "BO"
        .TextMatrix(1, 17) = "SQ"
        .TextMatrix(1, 18) = "ST"
        'VAN II
        .TextMatrix(1, 19) = "QH"
        .TextMatrix(1, 20) = "BO"
        .TextMatrix(1, 21) = "SQ"
        .TextMatrix(1, 22) = "ST"
        
        'PANABO
        .TextMatrix(1, 23) = "QH"
        .TextMatrix(1, 24) = "BO"
        .TextMatrix(1, 25) = "SQ"
        .TextMatrix(1, 26) = "ST"
        
        'POLOMOLOK
        .TextMatrix(1, 27) = "QH"
        .TextMatrix(1, 28) = "BO"
        .TextMatrix(1, 29) = "SQ"
        .TextMatrix(1, 30) = "ST"
        
        'SOBRECARY
        .TextMatrix(1, 31) = "QH"
        .TextMatrix(1, 32) = "BO"
        .TextMatrix(1, 33) = "SQ"
        .TextMatrix(1, 34) = "ST"
        
        'GENSAN
        .TextMatrix(1, 35) = "QH"
        .TextMatrix(1, 36) = "BO"
        .TextMatrix(1, 37) = "SQ"
        .TextMatrix(1, 38) = "ST"
        
        'MRC2RETAIL
        .TextMatrix(1, 39) = "QH"
        .TextMatrix(1, 40) = "BO"
        .TextMatrix(1, 41) = "SQ"
        .TextMatrix(1, 42) = "ST"
        
        'MARBEL
        .TextMatrix(1, 43) = "QH"
        .TextMatrix(1, 44) = "BO"
        .TextMatrix(1, 45) = "SQ"
        .TextMatrix(1, 46) = "ST"
        
        'STOTOMAS
        .TextMatrix(1, 47) = "QH"
        .TextMatrix(1, 48) = "BO"
        .TextMatrix(1, 49) = "SQ"
        .TextMatrix(1, 50) = "ST"
        
        'SAMAL
        .TextMatrix(1, 51) = "QH"
        .TextMatrix(1, 52) = "BO"
        .TextMatrix(1, 53) = "SQ"
        .TextMatrix(1, 54) = "ST"
        
        'VAN 3
        .TextMatrix(1, 55) = "QH"
        .TextMatrix(1, 56) = "BO"
        .TextMatrix(1, 57) = "SQ"
        .TextMatrix(1, 58) = "ST"
        
        'WHOLESALE
        .TextMatrix(1, 59) = "QH"
        .TextMatrix(1, 60) = "BO"
        .TextMatrix(1, 61) = "SQ"
        .TextMatrix(1, 62) = "ST"
        
        'MRC Cabantian
        .TextMatrix(1, 63) = "QH"
        .TextMatrix(1, 64) = "BO"
        .TextMatrix(1, 65) = "SQ"
        .TextMatrix(1, 66) = "ST"
        
        'Branch Total
        .TextMatrix(1, 67) = "QH"
        .TextMatrix(1, 68) = "BO"
        .TextMatrix(1, 69) = "SQ"
        .TextMatrix(1, 70) = "ST"

        'MRC WAREHOUSE
        .TextMatrix(1, 71) = "QH"
        .TextMatrix(1, 72) = "BO"
        .TextMatrix(1, 73) = "InTrn"
        
        'Purchase Order total
        .TextMatrix(1, 74) = "TTLSQ"
        .TextMatrix(1, 75) = cboMos.Text & "mSOQ"
        .TextMatrix(1, 76) = "ORDER"
        .TextMatrix(1, 77) = "CWOV"
        .TextMatrix(1, 78) = "APM"
        
        .MergeCells = flexMergeFree
        
        .MergeRow(0) = True
        For i = 1 To 78
            .MergeCol(i) = True
        Next i
        
        
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.13
        .ColWidth(2) = (.Width - 300) * 0.24
        For aCol = 3 To 6
            .ColWidth(aCol) = (.Width - 300) * 0.036
        Next aCol
        For aCol = 7 To 66  'YAMAHA TO WHOLESALE
            .ColWidth(aCol) = 0
        Next aCol
        For aCol = 67 To 74
            .ColWidth(aCol) = (.Width - 300) * 0.036
        Next aCol
        .ColWidth(75) = (.Width - 300) * 0.046
        .ColWidth(76) = (.Width - 300) * 0.04
        .ColWidth(77) = (.Width - 300) * 0.05
        .ColWidth(78) = (.Width - 300) * 0.05
        
        fScrol.Top = flxSPH.Top + 15
        fScrol.Height = .RowHeight(0)
        fScrol.Left = flxSPH.ColPos(6) + 80
    End With
End Sub

Sub ResizeCompleteFlxSP()
    Dim reCol As Integer
    With flxSP
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.13
        .ColWidth(2) = (.Width - 300) * 0.24
        
        .ColWidth(3) = (.Width - 300) * 0.036
        .ColWidth(4) = (.Width - 300) * 0.036
        .ColWidth(5) = (.Width - 300) * 0.036
        .ColWidth(6) = (.Width - 300) * 0.036
        fScrol.Left = flxSPH.ColPos(6) + 80
            
        For reCol = 7 To 78
            .ColWidth(reCol) = 0
            .ColAlignment(reCol) = 4
        Next reCol
        For reCol = 67 To 74
            .ColWidth(reCol) = (.Width - 300) * 0.036
        Next reCol
        .ColWidth(75) = (.Width - 300) * 0.046
        .ColWidth(76) = (.Width - 300) * 0.04
        .ColWidth(77) = (.Width - 300) * 0.05
        .ColWidth(78) = (.Width - 300) * 0.05
        .ColWidth(79) = 0
        .ColWidth(80) = 0
        
    End With
End Sub
Private Sub CollAlign()
    Dim nCols As Integer
     
    flxSP.ColAlignment(1) = 2
    flxSP.ColAlignment(2) = 2
    For nCols = 3 To 67
        flxSP.ColAlignment(nCols) = 4
    Next nCols
End Sub

Private Sub FillColorGrd()
    Dim ccol As Integer, cRow As Integer
    
    flxSP.Col = 3
    flxSP.Row = 0
    For ccol = 3 To 5
        flxSP.ColSel = ccol
        For cRow = 0 To flxSP.Rows - 1
            flxSP.RowSel = cRow
            flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = &HC0FFC0
        Next cRow
    Next ccol
End Sub


