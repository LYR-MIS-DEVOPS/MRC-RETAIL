VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BranchPO 
   Caption         =   "BRANCH P.O ORDER TAKING"
   ClientHeight    =   9225
   ClientLeft      =   2595
   ClientTop       =   465
   ClientWidth     =   13980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   13980
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtApprovedby 
      Height          =   285
      Left            =   9150
      TabIndex        =   47
      Tag             =   "XYF"
      Top             =   7890
      Width           =   4545
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   1830
      TabIndex        =   46
      Tag             =   "XYF"
      Top             =   7500
      Width           =   7365
   End
   Begin VB.TextBox txtPreparedby 
      Height          =   315
      Left            =   1830
      TabIndex        =   44
      Tag             =   "XYF"
      Top             =   7890
      Width           =   5595
   End
   Begin VB.TextBox txtOrder 
      Height          =   240
      Left            =   1395
      TabIndex        =   34
      Tag             =   "XYF"
      Top             =   9345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   135
      Top             =   9390
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   135
      Top             =   9870
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   735
      Top             =   9390
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   735
      Top             =   9870
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12270
      TabIndex        =   11
      Tag             =   "XYF"
      Text            =   "0.00"
      Top             =   7080
      Width           =   1470
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   7140
      Visible         =   0   'False
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   4605
      Left            =   300
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   2430
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   8123
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   $"BranchPO.frx":0000
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
   Begin VB.Frame Frame2 
      Caption         =   "Annotation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   825
      Left            =   300
      TabIndex        =   10
      Tag             =   "XYF"
      Top             =   8250
      Width           =   13410
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COWV = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   9870
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QT ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6555
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   510
         Width           =   360
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6570
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3675
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   450
         Width           =   420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SQ = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1005
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   450
         Width           =   405
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BO ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3675
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1005
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "COST w/Out VAT"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   10590
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Quantity Taking"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7050
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   510
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7065
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   180
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4170
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   450
         Width           =   1380
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sales Quantity"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1500
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   450
         Width           =   1020
      End
      Begin VB.Label Label6 
         Caption         =   "BACK-ORDER"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4170
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1455
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   180
         Width           =   1380
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxHead 
      Height          =   300
      Left            =   300
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   2160
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   529
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      FocusRect       =   2
      ScrollBars      =   0
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
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
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3570
      TabIndex        =   27
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "BranchPO.frx":00CF
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      HelpContextID   =   3
      Index           =   3
      Left            =   5355
      TabIndex        =   28
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
      Image           =   "BranchPO.frx":0669
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      HelpContextID   =   5
      Index           =   5
      Left            =   8640
      TabIndex        =   29
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "BranchPO.frx":0C03
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      HelpContextID   =   6
      Index           =   6
      Left            =   10440
      TabIndex        =   30
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
      Image           =   "BranchPO.frx":119D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      HelpContextID   =   6
      Index           =   7
      Left            =   12240
      TabIndex        =   31
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
      Image           =   "BranchPO.frx":1737
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1785
      TabIndex        =   32
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "BranchPO.frx":1CD1
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&AddNew"
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
      Image           =   "BranchPO.frx":226B
      cBack           =   -2147483633
   End
   Begin VB.Frame fmHeader 
      Height          =   1290
      Left            =   285
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   855
      Width           =   13455
      Begin VB.TextBox txtDateServe 
         Height          =   345
         Left            =   10980
         TabIndex        =   54
         Tag             =   "XYF"
         Top             =   360
         Width           =   1875
      End
      Begin VB.ComboBox cbodrtype 
         Height          =   315
         Left            =   7800
         TabIndex        =   52
         Tag             =   "XYF"
         Top             =   390
         Width           =   1815
      End
      Begin VB.TextBox txtDateCon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7800
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   750
         Width           =   1815
      End
      Begin VB.TextBox txtbranch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   375
         Width           =   2340
      End
      Begin VB.ComboBox cboMos 
         Height          =   315
         ItemData        =   "BranchPO.frx":2805
         Left            =   4050
         List            =   "BranchPO.frx":2807
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   750
         Width           =   1500
      End
      Begin lvButton.lvButtons_H cmdLoadPO 
         Height          =   330
         HelpContextID   =   4
         Left            =   4050
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         Caption         =   "&Add P.O"
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
         Image           =   "BranchPO.frx":2809
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblDateConf 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE SERVE"
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
         Left            =   9660
         TabIndex        =   53
         Tag             =   "XYF"
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label lblstatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   10980
         TabIndex        =   50
         Tag             =   "XYF"
         Top             =   780
         Width           =   1875
      End
      Begin VB.Label lblDateConf 
         Alignment       =   1  'Right Justify
         Caption         =   "STATUS"
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
         Left            =   9660
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblDateConf 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE CONFIRM"
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
         Left            =   6240
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MONTH/S (SQ)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER TYPE"
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
         Left            =   5880
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   450
         Width           =   1830
      End
      Begin VB.Label Label5 
         Caption         =   "BRANCH NAME"
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
         Left            =   420
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   435
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   51
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Update P.O"
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
      Image           =   "BranchPO.frx":2DA3
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAllPOM 
      Height          =   1875
      Left            =   5220
      TabIndex        =   56
      Top             =   9480
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   3307
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label3 
      Caption         =   "TOTAL "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9690
      TabIndex        =   55
      Tag             =   "XYF"
      Top             =   7140
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Index           =   2
      Left            =   7590
      TabIndex        =   48
      Tag             =   "XYF"
      Top             =   7905
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "REMARKS:"
      Height          =   255
      Left            =   300
      TabIndex        =   45
      Tag             =   "XYF"
      Top             =   7530
      Width           =   1515
   End
   Begin VB.Label Label12 
      Caption         =   "PREPARED BY:"
      Height          =   255
      Left            =   300
      TabIndex        =   43
      Tag             =   "XYF"
      Top             =   7920
      Width           =   1515
   End
   Begin VB.Label lblSQSERVE 
      Height          =   330
      Left            =   3300
      TabIndex        =   40
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblBID 
      Height          =   330
      Left            =   2685
      TabIndex        =   38
      Top             =   9360
      Visible         =   0   'False
      Width           =   420
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
      Left            =   2250
      TabIndex        =   35
      Top             =   9360
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizePos 
      Left            =   12060
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14100
      m_MDI_Height    =   9735
      OriginalBorder  =   2
      OriginalCaption =   27
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
      Left            =   315
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   540
      Width           =   2610
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      Height          =   255
      Index           =   0
      Left            =   8700
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   7140
      Width           =   495
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
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   570
      Width           =   2610
   End
   Begin VB.Menu mnuprinting 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnupick 
         Caption         =   "Order Picking"
      End
      Begin VB.Menu mnuconfirm 
         Caption         =   "Final Confirmation"
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "Order Taking"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Row"
      End
   End
End
Attribute VB_Name = "BranchPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddNew As Boolean                  'Boolean expression for Adding new records
Dim BtnIndx As Integer                 'Index for top Lacation of Button
Dim rsBrPO As New ADODB.Recordset
Dim pmos, myLastUpdt, Qry As String
Public mtrBName As String
Public isEdit As Boolean
Dim ActRow As Integer, ActCol As Integer, myOrderIDSQ As Long
Dim myOrderIDCof As Long

Private Sub cmdLoadPO_Click()
    If cboMos.Text = "" Then
        MsgBox ("Select SQ Range(1-6, All)"), vbExclamation, "Select"
        Exit Sub
    End If
    BranchPOSum.Show 1
End Sub
Private Sub cmdTrans_Click(Index As Integer)
Dim Ans As Integer
    Select Case Index
        Case 0
            Call ClearEntries
            BtnIndx = 5
            tmrFloat.Enabled = True
            cmdLoadPO.Enabled = True
            fmHeader.Enabled = True
            flxSP.Enabled = True
            txtDateServe = Format(Date, "mm/dd/yyyy")
            cbodrtype.Text = "(04) - STOCK"
            AddNew = True: isEdit = False
        Case 1
            If txtbranch.Text = "" Then
                MsgBox ("No Records to Edit"), vbExclamation, "Edit"
                Exit Sub
            End If
            BtnIndx = 5
            cmdLoadPO.Enabled = True
            tmrFloat.Enabled = True
            fmHeader.Enabled = True
            flxSP.Enabled = True
            isEdit = True
        Case 2
            FrmSummary.mySrc = "Branch Purchase Order"
            FrmSummary.Show
            Unload Me
        Case 3
            Call PopupMenu(mnuprinting)
        Case 4
            SPUPDATEPO.Show 1
        Case 5
            Unload Me
            MotoRep.tbMain.Visible = True
        Case 6
            If AddNew Then
                Ans = MsgBox("Save Order Taking Records?", vbQuestion + vbYesNo, "Save")
                If Ans = vbNo Then AddNew = False: Exit Sub
                Call SaveBranchOrder
                AddNew = False
            Else
                Ans = MsgBox("Update Order Taking Records?", vbQuestion + vbYesNo, "Save")
                If Ans = vbNo Then Exit Sub
                Call SaveBranchOrder
            End If
                BtnIndx = 7
                tmrFloatRev.Enabled = True
                fmHeader.Enabled = False
                flxSP.Enabled = False
        Case 7
            If lblstatus.Caption <> "" And Not AddNew Then
                Call ShowRecordOT(myOrderIDSQ)
            Else
                Call ClearEntries
            End If
            cmdLoadPO.Enabled = False
            BtnIndx = 7
            tmrFloatRev.Enabled = True
            AddNew = False
           
    End Select
End Sub
Sub ClearEntries()
    flxSP.Rows = 1
    flxSP.Rows = 2
    txtbranch.Text = ""
    'txtPONumber.Text = ""
    'cbOdrtype.Text = ""
    'txtPODate.Text = ""
    txtDateCon.Text = ""
    cboMos.ListIndex = 5
End Sub
Sub SaveBranchOrder()
    Dim pRw As Long, TotOrder As Long, TotBackOrder As Long, NoServe As Long, POTranstype As String
    Dim pos As Integer, TOTALQtyTK As Long, TotAmount As Currency, LastUpdtTk As String
    
        If flxSP.TextMatrix(1, 1) = "" Then
            MsgBox ("No Records were going to Update"), vbExclamation, "Empty"
            Exit Sub
        End If
             
             If AddNew Then
                'Qry = "SELECT MAX(spordertaking.SQSERVE) As SQLASTORDER From spordertaking Where spordertaking.BRANCHID= " & Val(lblBID.Caption) & " AND spordertaking.CONFIRMNO <> 'CANCELLED';"
                Qry = "SELECT MAX(spordertaking.SQSERVE) As SQLASTORDER From spordertaking;"
               
                NoServe = Val(GetFieldValue(Qry)) + 1
           
                LastUpdtTk = Format(Now, "yyyy-mm-dd hh:mm:ss")
                LYRcon.Execute "INSERT INTO spordertaking(BRANCHID, BRANCHNAME, MOS, DATESERVE, ORDERTYPE, QTYTAKING, BRANCHBO, TOTALAMOUNT, SQSERVE, REMARKS, PREPAREDBY, APPROVEDBY, LASTUPDATE, MYPC)VALUES(" & Val(lblBID.Caption) & ", " & _
                    "'" & txtbranch.Text & "', " & Val(cboMos) & ", '" & Format(txtDateServe.Text, "yyyy-mm-dd") & "', '" & cbodrtype.Text & "', " & TOTALQtyTK & ", " & TotBackOrder & ", " & TotAmount & ", " & NoServe & ", '" & txtRemarks.Text & "', " & _
                    "'" & txtPreparedby.Text & "', '" & txtApprovedby.Text & "', '" & LastUpdtTk & "', '" & ComputerName & "');"
                
                Qry = "SELECT spordertaking.ORDERIDSQ From spordertaking Where spordertaking.SQSERVE = " & NoServe & ";"
                ' AND spordertaking.MYPC = '" & ComputerName & "';"
                myOrderIDSQ = GetFieldValue(Qry)
                
                Qry = "SELECT MAX(spordrconfirm.ORDERID) AS LASTORDERID From spordrconfirm;"
                myOrderIDCof = Val(GetFieldValue(Qry)) + 1
           
                lblSQSERVE.Caption = NoServe
                POTranstype = "ORDER TEMP"
                
                For pRw = 1 To flxSP.Rows - 1
                    LYRcon.Execute "INSERT INTO spbrchconfirmdtl(ORDERID, SPAREPARTID, BRANCHID, " & _
                        "TRANSTYPE, MQTYONHAND, MBACKORDER, MQTYSALES, BQTYONHAND, BBACKORDER, BPORDER, ORDERTAKING, PRICE, REFERENCE, " & _
                        "UPDATEDBY)VALUES(" & myOrderIDCof & ", " & flxSP.TextMatrix(pRw, 13) & ", " & Val(lblBID.Caption) & ", '" & POTranstype & "', " & _
                        "" & Val(flxSP.TextMatrix(pRw, 3)) & ", " & Val(flxSP.TextMatrix(pRw, 4)) & ", " & Val(flxSP.TextMatrix(pRw, 5)) & ", " & _
                        "" & Val(flxSP.TextMatrix(pRw, 6)) & ", " & Val(flxSP.TextMatrix(pRw, 7)) & ", " & Val(flxSP.TextMatrix(pRw, 8)) & ", " & _
                        "" & Val(flxSP.TextMatrix(pRw, 9)) & ", " & Val(Replace(flxSP.TextMatrix(pRw, 10), ",", "")) & ", " & _
                        "'" & flxSP.TextMatrix(pRw, 12) & "', " & UserID & ");"
                Next pRw
                
                For pos = 1 To flxAllPOM.Rows - 1
                      
                    Qry = "SELECT SUM(spbrchconfirmdtl.BPORDER) As TOTALORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    TotOrder = Val(GetFieldValue(Qry))
    
                    Qry = "SELECT SUM(spbrchconfirmdtl.BBACKORDER) As TOTALBACKORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    TotBackOrder = Val(GetFieldValue(Qry))
                    
                    Qry = "SELECT SUM(spbrchconfirmdtl.ORDERTAKING) As TOTALORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    TOTALQtyTK = Val(GetFieldValue(Qry))
                    
                    Qry = "SELECT SUM(spbrchconfirmdtl.PRICE) As TOTALORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    TotAmount = Val(GetFieldValue(Qry))
                    
                    myLastUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
                    LYRcon.Execute "INSERT INTO spordrconfirm(ORDERID, BRANCHID, BRANCHNAME, REFERENCE, ORDERTYPE, QTYORDER, BACKORDER, TOTALAMOUNTPO, SQSERVE, LASTUPDATE, MYPC)VALUES(" & myOrderIDCof & ", " & Val(lblBID.Caption) & ", " & _
                        "'" & txtbranch.Text & "', '" & flxAllPOM.TextMatrix(pos, 0) & "', '" & cbodrtype.Text & "', " & TotOrder & ", " & TotBackOrder & ", " & Replace(txtTotalAmount.Text, ",", "") & ", " & _
                        "" & NoServe & ", '" & myLastUpdt & "', '" & ComputerName & "');"
                
                Next pos
                
                MsgBox ("Purchase Order Successfully Save"), vbInformation
            Else
                    
                    LastUpdtTk = Format(Now, "yyyy-mm-dd hh:mm:ss")
                   
                    LYRcon.Execute "UPDATE spordertaking SET BRANCHID = " & Val(lblBID.Caption) & ", BRANCHNAME = '" & txtbranch.Text & "', MOS = " & Val(cboMos) & ", " & _
                        "DATESERVE = '" & Format(txtDateServe.Text, "yyyy-mm-dd") & "', ORDERTYPE = '" & cbodrtype.Text & "', QTYTAKING = " & TOTALQtyTK & ", " & _
                        "BRANCHBO = " & TotBackOrder & ", TOTALAMOUNT = " & TotAmount & ", SQSERVE = " & Val(lblSQSERVE.Caption) & ", PREPAREDBY = '" & txtPreparedby.Text & "', " & _
                        "REMARKS = '" & txtRemarks.Text & "', APPROVEDBY = '" & txtApprovedby.Text & "', LASTUPDATE = '" & LastUpdtTk & "' Where spordertaking.ORDERIDSQ = " & myOrderIDSQ & ";"
                        'lblSQSERVE.Caption = NoServe
                        POTranstype = "ORDER TEMP"
                    
                    
                    LYRcon.Execute "DELETE from spbrchconfirmdtl " & _
                        "Where spbrchconfirmdtl.ORDERID = " & myOrderIDCof & " AND spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & ";"
                    For pRw = 1 To flxSP.Rows - 1
                     
                     LYRcon.Execute "INSERT INTO spbrchconfirmdtl(ORDERID, SPAREPARTID, BRANCHID, " & _
                        "TRANSTYPE, MQTYONHAND, MBACKORDER, MQTYSALES, BQTYONHAND, BBACKORDER, BPORDER, ORDERTAKING, PRICE, REFERENCE, " & _
                        "UPDATEDBY)VALUES(" & myOrderIDCof & ", " & flxSP.TextMatrix(pRw, 13) & ", " & Val(lblBID.Caption) & ", '" & POTranstype & "', " & _
                        "" & Val(flxSP.TextMatrix(pRw, 3)) & ", " & Val(flxSP.TextMatrix(pRw, 4)) & ", " & Val(flxSP.TextMatrix(pRw, 5)) & ", " & _
                        "" & Val(flxSP.TextMatrix(pRw, 6)) & ", " & Val(flxSP.TextMatrix(pRw, 7)) & ", " & Val(flxSP.TextMatrix(pRw, 8)) & ", " & _
                        "" & Val(flxSP.TextMatrix(pRw, 9)) & ", " & Val(Replace(flxSP.TextMatrix(pRw, 10), ",", "")) & ", " & _
                        "'" & flxSP.TextMatrix(pRw, 12) & "', " & UserID & ");"
                     Next pRw
                    
                     For pos = 1 To flxAllPOM.Rows - 1
                        
                        Qry = "SELECT SUM(spbrchconfirmdtl.BPORDER) As TOTALORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & ";"
                        TotOrder = Val(GetFieldValue(Qry))
    
                        Qry = "SELECT SUM(spbrchconfirmdtl.BBACKORDER) As TOTALBACKORDER From spbrchconfirmdtl " & _
                            " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & ";"
                        TotBackOrder = Val(GetFieldValue(Qry))
                        
                        Qry = "SELECT SUM(spbrchconfirmdtl.ORDERTAKING) As TOTALORDER From spbrchconfirmdtl " & _
                            " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & ";"
                        TOTALQtyTK = Val(GetFieldValue(Qry))
                        
                        Qry = "SELECT SUM(spbrchconfirmdtl.PRICE) As TOTALORDER From spbrchconfirmdtl " & _
                            " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & ";"
                        TotAmount = Val(GetFieldValue(Qry))
     
                        myLastUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
                        LYRcon.Execute "Update spordrconfirm SET BRANCHID = " & Val(lblBID.Caption) & ", BRANCHNAME = '" & txtbranch.Text & "', " & _
                            "ORDERTYPE = '" & cbodrtype.Text & "', QTYORDER = " & TotOrder & ", BACKORDER = " & TotBackOrder & ", TOTALAMOUNTPO = " & TotAmount & ", " & _
                            "SQSERVE = " & Val(lblSQSERVE.Caption) & ", LASTUPDATE = '" & myLastUpdt & "', MYPC = '" & ComputerName & "' " & _
                            "Where spordrconfirm.SQSERVE = " & myOrderIDSQ & " AND spordrconfirm.REFERENCE = '" & flxAllPOM.TextMatrix(pos, 0) & "' AND spordrconfirm.BRANCHID = " & Val(lblBID.Caption) & ";"
                    
                    Next pos
                MsgBox ("Purchase Order Successfully Updated"), vbInformation
             End If
End Sub
Private Sub flxSP_Click()
    ActRow = flxSP.Row
    ActCol = flxSP.Col
    If ActCol <> 9 Then Exit Sub
    
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
End Sub
Private Sub flxSP_KeyPress(KeyAscii As Integer)
    Call flxSP_Click
End Sub
Private Sub flxSP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuOrder
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSP_Click
    fmHeader.Enabled = False
End Sub
Private Sub Form_Load()
    Dim MySupp As String, m As Integer
    
    MotoRep.tbMain.Visible = False
    For m = 1 To 6
        cboMos.AddItem m
    Next m
    
    cboMos.AddItem "All"
    cboMos.ListIndex = 5
    lblDate.Caption = DateAdd("m", -6, Date) & " - " & Date
    pmos = -6
    
    cbodrtype.AddItem "(01) - URGENT"
    cbodrtype.AddItem "(04) - STOCK"
    
    fmHeader.Enabled = False
    flxSP.Enabled = False
End Sub
Private Sub mnuconfirm_Click()
    Dim Qry As String
    Dim Ans As Integer, pRw As Long, ref As Long
        
        If lblstatus.Caption = "" Then
            MsgBox ("No Records to Print"), vbExclamation, "No Record"
            Exit Sub
        End If
        Qry = "SELECT spordertaking.PRINTPICK from spordertaking Where spordertaking.BRANCHNAME = '" & txtbranch & "' AND spordertaking.SQSERVE = " & Val(lblSQSERVE.Caption) & ";"
        If GetFieldValue(Qry) = 0 Then
            Ans = MsgBox("Please prepare Print-out for picking", vbQuestion, "Order Picking")
            Exit Sub
        End If
        Qry = "SELECT spordertaking.DATECONFIRM from spordertaking Where spordertaking.BRANCHNAME = '" & txtbranch & "' AND spordertaking.SQSERVE = " & Val(lblSQSERVE.Caption) & ";"
        If GetFieldValue(Qry) = Null Or GetFieldValue(Qry) = "" Then
            Ans = MsgBox("Print Final Confirmation", vbQuestion + vbYesNo, "Order Picking")
            If Ans = vbNo Then Exit Sub
        End If
           LYRcon.Execute "DELETE from spbrchconfirmdtl " & _
                            "Where spbrchconfirmdtl.ORDERID = " & myOrderIDCof & " AND spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERTAKING=0;"
            For ref = 1 To flxAllPOM.Rows - 1
                    Qry = "SELECT SUM(spbrchconfirmdtl.BBACKORDER) As TOTALBACKORDER From spbrchconfirmdtl " & _
                        " where spbrchconfirmdtl.REFERENCE = '" & flxAllPOM.TextMatrix(ref, 0) & "' AND  spbrchconfirmdtl.BRANCHID= " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.ORDERID='" & myOrderIDCof & "';"
                    If Val(GetFieldValue(Qry)) = 0 Then
                        LYRcon.Execute "Update sporderbranch SET ALLOCATED = 1 Where sporderbranch.REFERENCE =  '" & flxAllPOM.TextMatrix(ref, 0) & "' AND sporderbranch.BRANCHID = " & Val(lblBID.Caption) & ";"
                    End If
            Next ref
        
        DoEvents
        
        PRNmc.MySource = "SPBPOACONFIRM"
        PRNmc.QryPrn = "Select DISTINCT spbrchconfirmdtl.ORDERTAKING, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spbrchconfirmdtl.REFERENCE, spareparts.COST, spbrchconfirmdtl.ORDERTAKING * spareparts.COST AS TOTALAMOUNT " & _
            "From spbrchconfirmdtl INNER JOIN spareparts " & _
                "ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) INNER JOIN spordrconfirm " & _
                "ON (spbrchconfirmdtl.ORDERID = spordrconfirm.ORDERID) " & _
            "Where spbrchconfirmdtl.BRANCHID =  " & Val(lblBID.Caption) & " AND spordrconfirm.SQSERVE = " & Val(lblSQSERVE.Caption) & " ORDER BY spareparts.PARTNUMBER;"
        
        PRNmc.ParamValue1 = SysBranch
        PRNmc.ParamValue2 = SysAddress
        PRNmc.ParamValue3 = SysTel
        PRNmc.ParamValue4 = txtbranch.Text
        Qry = "SELECT branch.ADDRESS From branch Where branch.BRANCHID = " & Val(lblBID) & ";"
        PRNmc.ParamValue5 = GetFieldValue(Qry)
        PRNmc.ParamValue6 = txtDateServe.Text
        PRNmc.ParamValue7 = lblSQSERVE.Caption
        PRNmc.ParamValue8 = txtPreparedby.Text
        PRNmc.ParamValue9 = CStr(txtDateCon.Text)
        PRNmc.ParamValue10 = "JAMIE CARLO N. ANG"
        
        PRNmc.Show
 End Sub

Private Sub mnuDelete_Click()
Dim Ans As Integer
    Ans = MsgBox("Delete Current Row record?", vbYesNo + vbQuestion, "Delete Row")
    If Ans = vbYes Then
        If flxSP.Rows = 2 Then
            flxSP.Rows = 1
            flxSP.Rows = 2
        Else
            flxSP.RemoveItem (flxSP.Row)
        End If
    End If
End Sub
Private Sub mnupick_Click()
    Dim Qry As String
    Dim Ans As Integer
        
        If lblstatus.Caption = "" Then
            MsgBox ("View Records to Print"), vbExclamation, "No Record"
            Exit Sub
        End If
        'Qry = "SELECT spordrconfirm.PRINTPICK from spordrconfirm Where spordrconfirm.BRANCHNAME = '" & txtbranch & "' AND  spordrconfirm.REFERENCE = '" & txtPONumber.Text & "' AND  spordrconfirm.SQSERVE = " & Val(lblSQSERVE.Caption) & ";"
        Qry = "SELECT spordertaking.PRINTPICK from spordertaking Where spordertaking.BRANCHNAME = '" & txtbranch & "' AND  spordertaking.SQSERVE = " & Val(lblSQSERVE.Caption) & ";"
         
        If GetFieldValue(Qry) = 0 Then
            Ans = MsgBox("Do you want to Print P.O for Order Picking", vbQuestion + vbYesNo, "Order Picking")
            If Ans = vbNo Then Exit Sub
            mnuconfirm.Enabled = True
        End If
        'One Reference Id
        
        PRNmc.MySource = "SPBPO"
        If flxAllPOM.Rows = 2 Then
            PRNmc.QryPrn = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.LOCATION, spbrchconfirmdtl.REFERENCE, spbrchconfirmdtl.BBACKORDER " & _
                " From spbrchconfirmdtl " & _
                " INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                " INNER JOIN spordertaking ON (spbrchconfirmdtl.BRANCHID = spordertaking.BRANCHID) " & _
                " AND (spbrchconfirmdtl.ORDERID = spordertaking.SQSERVE) " & _
                " Where spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & "  AND " & _
                " spordertaking.ORDERIDSQ = " & myOrderIDSQ & " AND spbrchconfirmdtl.BBACKORDER > 0 Order By spareparts.PARTNUMBER;"
        Else
            PRNmc.QryPrn = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.LOCATION, " & _
                " spbrchconfirmdtl.REFERENCE, spbrchconfirmdtl.BBACKORDER " & _
                "From spbrchconfirmdtl " & _
                "INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                "INNER JOIN spordertaking ON (spbrchconfirmdtl.BRANCHID = spordertaking.BRANCHID) " & _
                "AND (spbrchconfirmdtl.ORDERID = spordertaking.SQSERVE) " & _
                "Where spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & " AND spbrchconfirmdtl.BBACKORDER > 0 Order By SPAREPARTS.PartNumber;"
        
        End If
        'PRNmc.MySource = "SPBPO"
        'PRNmc.QryPrn = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.LOCATION, spbrchconfirmdtl.REFERENCE, spbrchconfirmdtl.BBACKORDER " & _
        '    "From spbrchconfirmdtl " & _
        '    "INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
        '    "INNER JOIN spordertaking ON (spbrchconfirmdtl.BRANCHID = spordertaking.BRANCHID) " & _
        '    "Where spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & " AND spordertaking.ORDERIDSQ = " & myOrderIDSQ & " AND spbrchconfirmdtl.BBACKORDER>0 ORDER BY spareparts.PARTNUMBER;"
        
        'PRNmc.QryPrn = "SELECT DISTINCT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.LOCATION, spbrchconfirmdtl.REFERENCE, " & _
            "spbrchconfirmdtl.BBACKORDER From spbrchconfirmdtl " & _
            "INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
            "INNER JOIN spordertaking ON (spbrchconfirmdtl.BRANCHID = spordertaking.BRANCHID) " & _
            "AND (spbrchconfirmdtl.ORDERID = spordertaking.ORDERIDSQ) " & _
            "Where spbrchconfirmdtl.BRANCHID = " & Val(lblBID.Caption) & " AND spordertaking.ORDERIDSQ = " & myOrderIDSQ & ";"
       
        PRNmc.ParamValue1 = SysBranch
        PRNmc.ParamValue2 = SysAddress
        PRNmc.ParamValue3 = SysTel
        PRNmc.ParamValue4 = txtbranch.Text
        Qry = "SELECT branch.ADDRESS From branch Where branch.BRANCHID = '" & Val(lblBID) & "';"
        PRNmc.ParamValue5 = GetFieldValue(Qry)
        PRNmc.ParamValue6 = txtDateServe.Text
        PRNmc.ParamValue7 = lblSQSERVE.Caption
        PRNmc.ParamValue8 = txtPreparedby.Text
        PRNmc.Show
End Sub

Private Sub RXResizePos_ResizeComplete()
    Call ResizeCompleteHead
End Sub
Private Sub tmrDrop_Timer() 'Timer for drop position of buton
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 8 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub
Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then
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
            BtnIndx = 6 'Index of Button Save
            cmdTrans(6).Top = -375
            cmdTrans(6).Left = 0
            cmdTrans(6).Visible = True
            cmdTrans(7).Top = -375
            cmdTrans(7).Left = cmdTrans(4).Width
            cmdTrans(7).Visible = True
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
            For i = 0 To 5
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub
Private Sub ResizeCompleteHead()
Dim AdCol, aCol As Integer, aCel As Integer, reAln As Integer, BrWidth As Long
    
    With flxSP
        For aCol = 0 To 2
            msflxHead.ColAlignment(aCol) = 4
        Next aCol
        msflxHead.TextMatrix(0, 0) = "S P A R E P A R T S"
        msflxHead.TextMatrix(0, 1) = "MOTO-R MAIN"
        msflxHead.TextMatrix(0, 2) = "BRANCH"
        
        
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - 400) * 0.15
        .ColWidth(2) = (.Width - 400) * 0.24
        .ColWidth(3) = (.Width - 400) * 0.04
        .ColWidth(4) = (.Width - 400) * 0.04
        
        .ColWidth(5) = (.Width - 400) * 0.04
        .ColWidth(6) = (.Width - 400) * 0.04
        .ColWidth(7) = (.Width - 400) * 0.04
        .ColWidth(8) = (.Width - 400) * 0.04
        .ColWidth(9) = (.Width - 400) * 0.04
        .ColWidth(10) = (.Width - 400) * 0.07
        .ColWidth(11) = (.Width - 400) * 0.08
        .ColWidth(12) = (.Width - 400) * 0.173
        .ColWidth(13) = 0
        msflxHead.ColWidth(0) = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 10
        msflxHead.ColWidth(1) = .ColWidth(3) * 3
       
        For AdCol = 6 To 11
            BrWidth = BrWidth + .ColWidth(AdCol)
        Next AdCol
        msflxHead.ColWidth(2) = BrWidth + .ColWidth(12)
        fmHeader.Width = flxSP.Width
        txtTotalAmount.Width = flxSP.ColWidth(11)
        txtTotalAmount.Left = txtTotalAmount.Left - flxSP.ColWidth(12)
        Label3.Left = txtTotalAmount.Left - txtTotalAmount.Width
        
        
    End With
End Sub
Sub ShowRecord(BId As Integer)  'P.O from Branch
    Dim rsPdn As New ADODB.Recordset
    Dim Qry As String, TOtPPO As Long, BOrderBal As Long, nR As Long
    Dim nRow As Long, fl As Integer, myW As String, myW1 As String, ASQSERVE As Long
    Dim r As Integer, MyRefNo As String
    
    If BranchPOSum.chkPO.Value = 1 Then
        
        flxAllPOM.Clear
        flxAllPOM.Rows = 1
        BranchPOSum.flxAllPO.FixedRows = 1
        
        For r = 1 To BranchPOSum.flxAllPO.Rows - 1
             
            Qry = "SELECT * From spordertaking INNER JOIN spordrconfirm ON (spordertaking.SQSERVE = spordrconfirm.ORDERID) " & _
                  " Where spordrconfirm.REFERENCE = '" & BranchPOSum.flxAllPO.TextMatrix(r, 0) & "' AND spordertaking.FINALQT = 0;"
            If rsPdn.State = adStateOpen Then rsPdn.Close
            rsPdn.Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
            If rsPdn.RecordCount > 0 Then
                PdnPO = PdnPO & BranchPOSum.flxAllPO.TextMatrix(r, 0) & " "
            End If
            Qry = "SELECT sporderbranch.ALLOCATED From sporderbranch Where sporderbranch.REFERENCE = '" & BranchPOSum.flxAllPO.TextMatrix(r, 0) & "' AND sporderbranch.BRANCHID = " & BId & ";"
            If Val(GetFieldValue(Qry)) = 0 Then
                flxAllPOM.AddItem BranchPOSum.flxAllPO.TextMatrix(r, 0)
            End If
        Next r
        
        If PdnPO <> "" Then
            MsgBox ("Reference No: " & PdnPO & " are Already Added for Order Taking..Please Print for Final Confirmation"), vbExclamation, "PENDING"
            flxAllPOM.Clear
            flxAllPOM.Rows = 2
            flxAllPOM.FixedRows = 1
            Exit Sub
        End If
        flxAllPOM.FixedRows = 1
    Else
         MyRefNo = BranchPOSum.cboRefNo.Text
         flxAllPOM.Clear
         flxAllPOM.Rows = 1
         
         flxAllPOM.AddItem MyRefNo
         flxAllPOM.FixedRows = 1
    End If
    cmdLoadPO.Enabled = False
   
    myW1 = ""
    If BranchPOSum.chkPO.Value = 0 Then
        myW1 = "AND sporderbranch.REFERENCE = '" & MyRefNo & "'"
    Else
        For i = 1 To flxAllPOM.Rows - 1
            If i = 1 Then
                myW1 = myW1 & " AND sporderbranch.REFERENCE = '" & flxAllPOM.TextMatrix(i, 0) & "'"
            Else
                myW1 = myW1 & " OR sporderbranch.REFERENCE = '" & flxAllPOM.TextMatrix(i, 0) & "'"
            End If
        Next i
    End If
    myW = myW1
    
    Qry = "SELECT sporderbranch.REFERENCE, sporderbranch.BRANCHNAME, sporderbranch.BRANCHID, sporderbranch.ORDERTYPE, sporderbranch.ORDERDATE " & _
          "From sporderbranch Where sporderbranch.BRANCHID = " & BId & " " & myW & " ORDER BY sporderbranch.REFERENCE;"
    
    With rsBrPO
         If .State = adStateOpen Then .Close
         .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
             Call ClearEntries
             If .RecordCount = 0 Then Exit Sub: .Close
                 
                 lblBID.Caption = rsBrPO!BRANCHID
                 txtbranch.Text = rsBrPO!BRANCHNAME
                 'txtPONumber.Text = rsBrPO!REFERENCE
                 'txtOdrtype.Text = rsBrPO!ORDERTYPE
                 'txtPODate.Text = rsBrPO!ORDERDATE
                 lblstatus.Caption = "PENDING"
                 .Close
                myW = ""
                myW1 = ""
                If BranchPOSum.chkPO.Value = 0 Then
                        myW1 = "AND spordrconfirm.REFERENCE = '" & MyRefNo & "'"
                Else
                    For i = 0 To flxAllPOM.Rows - 1
                        If i = 1 Then
                            myW1 = myW1 & " AND spordrconfirm.REFERENCE = '" & flxAllPOM.TextMatrix(i, 0) & "'"
                        Else
                            myW1 = myW1 & " OR spordrconfirm.REFERENCE = '" & flxAllPOM.TextMatrix(i, 0) & "'"
                        End If
                    Next i
                End If
                myW = myW1
                 
            'Qry = "SELECT SPORDERTAKING.SQSERVE  From SPORDERTAKING Where SPORDERTAKING.BRANCHID = " & Val(lblBID.Caption) & " AND SPORDERTAKING.FINALQT = 0;"
            'ASQSERVE = Val(GetFieldValue(Qry))
            
            'Qry = "SELECT spordrconfirm.ORDERID From spordrconfirm Where spordrconfirm.SQSERVE = " & ASQSERVE & ";"
           
            'If .State = adStateOpen Then .Close
            '.Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
            'If .RecordCount > 0 Then
            '        MsgBox ("Please Print Previous Order Taking"), vbExclamation, "   Not Confirm P.O"
            '        .Close
            '         Call ClearEntries
            '        Exit Sub
            'End If
                With BranchPOSum.flxSP
                    fl = 0
                    For nRow = 1 To .Rows - 1
                        If Val(.TextMatrix(nRow, 6)) > 0 Then
                                fl = fl + 1
                                flxSP.Rows = fl + 1
                                flxSP.TextMatrix(fl, 8) = .TextMatrix(nRow, 4)
                                flxSP.TextMatrix(fl, 9) = ""
                                flxSP.TextMatrix(fl, 13) = .TextMatrix(nRow, 9)
                                flxSP.TextMatrix(fl, 7) = .TextMatrix(nRow, 6)
                                flxSP.TextMatrix(fl, 12) = .TextMatrix(nRow, 3)
                                
                        End If
                    Next nRow
                   
                End With
                 If flxSP.Rows = 2 And flxSP.TextMatrix(1, 1) = "" Then flxAllPOM.Clear: flxAllPOM.Rows = 2: flxAllPOM.FixedRows = 1: Exit Sub
                If .State = adStateOpen Then .Close
                For nRow = 1 To flxSP.Rows - 1
                        Qry = "SELECT spbranchdtl.SPAREPARTID, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.BACKORDER, spbranchdtl.BQTYONHAND, " & _
                              "spareparts.cost From spbranchdtl INNER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                              "Where spareparts.SPAREPARTID =  " & Val(flxSP.TextMatrix(nRow, 13)) & " AND spbranchdtl.REFERENCE ='" & flxSP.TextMatrix(nRow, 12) & "' AND spbranchdtl.BRANCHNAME = '" & mtrBName & "';"
                        If .State = adStateOpen Then .Close
                        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                        If .RecordCount > 0 Then
                            flxSP.TextMatrix(nRow, 1) = !PartNumber
                            flxSP.TextMatrix(nRow, 2) = !Description
                            flxSP.TextMatrix(nRow, 3) = !QtyOnHand
                            flxSP.TextMatrix(nRow, 4) = !BACKORDER
                            flxSP.TextMatrix(nRow, 5) = Val(GetSalesQtyMain(flxSP.TextMatrix(nRow, 13)))
                            flxSP.TextMatrix(nRow, 6) = !BQTYONHAND
                            flxSP.TextMatrix(nRow, 10) = Format(!Cost, "#,###,###.#0")
                            flxSP.TextMatrix(nRow, 11) = "0.00"
                        End If
                 Next nRow
                 cmdTrans(1).Enabled = True
                 cmdTrans(3).Enabled = True
                 
                 .Close
    End With
End Sub
Sub ShowRecordOT(SVRid As Long) 'ORDERIDSQ

Dim Qry As String, nRow As Long, ref As Integer
Dim rsPO As New ADODB.Recordset, rsRef As New ADODB.Recordset
    
    flxSP.Rows = 1: flxSP.Rows = 2
    fmHeader.Enabled = False: flxSP.Enabled = False
    myOrderIDSQ = SVRid
    With rsPO
        
        Qry = "SELECT spordertaking.ORDERIDSQ, spordertaking.BRANCHID, spordertaking.BRANCHNAME, spordertaking.MOS, spordertaking.DATESERVE, spordertaking.SQSERVE , spordertaking.REMARKS, spordertaking.PREPAREDBY, " & _
                "spordertaking.APPROVEDBY, spordertaking.PRINTPICK, spordertaking.DATECONFIRM, spordertaking.FINALQT " & _
                "From spordertaking WHERE spordertaking.SQSERVE = " & myOrderIDSQ & ";"
        
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
        If .RecordCount > 0 Then
            .MoveFirst
            lblBID.Caption = !BRANCHID
            txtbranch.Text = !BRANCHNAME
            cboMos.Text = !mos
            txtDateServe.Text = !DATESERVE
            txtDateCon.Text = IIf(IsNull(!DATECONFIRM), "", !DATECONFIRM)
            lblSQSERVE.Caption = !SQSERVE
            myOrderIDSQ = !ORDERIDSQ
            cbodrtype.Text = "(04) - STOCK"
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPreparedby.Text = IIf(IsNull(!PreparedBy), "", !PreparedBy)
            txtApprovedby.Text = IIf(IsNull(!ApprovedBy), "", !ApprovedBy)
            If !PRINTPICK >= 1 And Not IsNull(!DATECONFIRM) Then
                cmdTrans(1).Enabled = False
                mnupick.Enabled = False
                mnuconfirm.Enabled = True
            ElseIf !PRINTPICK >= 1 And IsNull(!DATECONFIRM) Then
                cmdTrans(1).Enabled = True
                mnupick.Enabled = True
                mnuconfirm.Enabled = True
            ElseIf !PRINTPICK = 0 Then
                cmdTrans(1).Enabled = True
                mnupick.Enabled = True
                mnuconfirm.Enabled = False
            End If
            If .Fields("FINALQT") = 0 Then
                lblstatus.Caption = "PENDING"
            ElseIf .Fields("FINALQT") = 1 Then
                lblstatus.Caption = "FINAL"
            ElseIf .Fields("FINALQT") = 2 Then
                lblstatus.Caption = "CANCELLED"
            End If
            .Close
            
        End If
        
        Qry = "SELECT spordrconfirm.ORDERID FROM spordrconfirm WHERE spordrconfirm.SQSERVE = " & Val(lblSQSERVE.Caption) & " AND  spordrconfirm.BRANCHID = " & Val(lblBID.Caption) & ";"
        myOrderIDCof = Val(GetFieldValue(Qry))
        
        Qry = "SELECT spbrchconfirmdtl.REFERENCE, spbrchconfirmdtl.SPAREPARTID, spbrchconfirmdtl.BPORDER, spbrchconfirmdtl.BBACKORDER, spbrchconfirmdtl.ORDERTAKING, spbrchconfirmdtl.PRICE, spbrchconfirmdtl.REMARKS " & _
               "From spbrchconfirmdtl Where spbrchconfirmdtl.ORDERID =" & myOrderIDCof & ";"
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
             If .RecordCount = 0 Then Exit Sub: .Close
                 flxSP.Rows = .RecordCount + 1
                  .MoveFirst
                    For nRow = 1 To .RecordCount
                          flxSP.TextMatrix(nRow, 8) = rsPO(2).Value
                          flxSP.TextMatrix(nRow, 9) = IIf(IsNull(!ORDERTAKING) Or !ORDERTAKING = 0, "", !ORDERTAKING)
                          flxSP.TextMatrix(nRow, 7) = IIf(IsNull(!BBACKORDER), 0, !BBACKORDER)
                          flxSP.TextMatrix(nRow, 12) = IIf(IsNull(!REFERENCE), "", !REFERENCE)
                          flxSP.TextMatrix(nRow, 13) = !SPAREPARTID
                         .MoveNext
                     Next nRow
                    .Close
                    Qry = "SELECT spordrconfirm.REFERENCE FROM spordrconfirm WHERE spordrconfirm.SQSERVE = " & Val(lblSQSERVE.Caption) & " AND spordrconfirm.BRANCHID = " & Val(lblBID.Caption) & ";"
                    If rsRef.State = adStateOpen Then rsRef.Close
                    rsRef.Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                    rsRef.MoveFirst
                     flxAllPOM.Clear
                    Do While Not rsRef.EOF
                        flxAllPOM.AddItem rsRef!REFERENCE
                        rsRef.MoveNext
                    Loop
                        flxAllPOM.Rows = 2
                        flxAllPOM.FixedRows = 1
                    
                     For nRow = 1 To flxSP.Rows - 1
                        Qry = "SELECT spbrchconfirmdtl.SPAREPARTID, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.BACKORDER, spbrchconfirmdtl.BQTYONHAND, " & _
                              "spareparts.COST From spbrchconfirmdtl INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                              "Where spareparts.SPAREPARTID =  " & flxSP.TextMatrix(nRow, 13) & " AND spbrchconfirmdtl.ORDERID=" & myOrderIDCof & " ORDER BY spareparts.PARTNUMBER;"
                        If .State = adStateOpen Then .Close
                        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                        flxSP.TextMatrix(nRow, 0) = nRow
                        flxSP.TextMatrix(nRow, 1) = !PartNumber
                        flxSP.TextMatrix(nRow, 2) = !Description
                        flxSP.TextMatrix(nRow, 3) = !QtyOnHand
                        flxSP.TextMatrix(nRow, 4) = !BACKORDER
                        flxSP.TextMatrix(nRow, 5) = Val(GetSalesQtyMain(flxSP.TextMatrix(nRow, 13)))
                        flxSP.TextMatrix(nRow, 6) = !BQTYONHAND
                        flxSP.TextMatrix(nRow, 10) = Format(!Cost, "#,###,###.#0")
                    Next nRow
                    Call ComputeTotalView
        .Close
    End With
End Sub
Private Sub cbomos_Change()
    cbomos_Click
End Sub
Private Sub cbomos_Click()
Dim mos As Integer
    mos = Val(cboMos.Text)
    pmos = "-" & mos
    lblDate.Caption = DateAdd("m", pmos, Date) & "  TO  " & Date
    If cboMos.Text = "All" Then
          lblDate.Caption = "12/19/2007" & "  TO  " & Date
    End If
End Sub
Function GetSalesQtyMain(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", pmos, "12/19/2007"), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
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
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", pmos, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQtyMain = MySales - MyRet
End Function
Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If flxSP.Col = 9 Then
            If ValidateEntry = True Then
                txtOrder.Text = ""
                txtOrder.Visible = False
                flxSP.SetFocus
                flxSP.Row = ActRow
                flxSP.Col = 9
                txtOrder.ZOrder
                Exit Sub
            End If
                txtOrder.Text = flxSP.TextMatrix(ActRow, 9)
                flxSP.SetFocus
                flxSP.Row = ActRow
                flxSP.Col = 9
                txtOrder.Visible = False
        End If
    End If
End Sub
Private Function ValidateEntry() As Boolean
    With txtOrder
        If txtOrder.Text = "" Then Exit Function
        'If flxsp.Col = 9 Then
            If Val(.Text) > flxSP.TextMatrix(ActRow, 7) Or Val(.Text) > flxSP.TextMatrix(ActRow, 8) Then
                MsgBox ("Order Taking not Greater Than the Backorder, Ordered parts"), vbExclamation, "Order"
                .SelStart = 0
                .SelLength = Len(txtOrder.Text)
                .ZOrder
                .SetFocus
                ValidateEntry = True
            ElseIf Val(.Text) > flxSP.TextMatrix(ActRow, 8) Then
                MsgBox ("Order Taking not Greater Than the Ordered parts"), vbExclamation, "Order"
                .SelStart = 0
                .SelLength = Len(txtOrder.Text)
                .ZOrder
                .SetFocus
                ValidateEntry = True
            ElseIf Val(.Text) > flxSP.TextMatrix(ActRow, 3) Then
                MsgBox ("Quantity Order Taking not Greater Than Available stock"), vbExclamation, "Order"
                .SelStart = 0
                .SelLength = Len(txtOrder.Text)
                .ZOrder
                .SetFocus
                ValidateEntry = True
            Else
                ValidateEntry = False
            End If
            'ElseIf flxsp.Col = 8 Then
            '    If Val(.Text) > Val(flxsp.TextMatrix(ActRow, 3)) Then
            '        MsgBox ("Quantity Order not Greater Than Available stock"), vbExclamation, "Order"
            '        .SelStart = 0
            '        .SelLength = Len(txtOrder.Text)
            '        .ZOrder
            '        .SetFocus
            '        ValidateEntry = True
            '    End If
            'End If
     End With
End Function
Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
            flxSP.SetFocus
            flxSP.Row = ActRow
            
            flxSP.Col = 9
    End If
End Sub
Private Sub txtOrder_LostFocus()
    With flxSP
        If ValidateEntry = True Then Exit Sub
        
        If .Col = 9 Then
            .TextMatrix(ActRow, 9) = txtOrder.Text
             txtOrder.Visible = False
             Call ComputeTotal
             If Val(.TextMatrix(ActRow, 9)) > 0 Then
                 .Row = ActRow
                 .Col = 9
                 '.CellFontName = "WINGDINGS"
                 '.CellFontSize = 12
                 '.Text = lblChk.Caption
                 '.CellAlignment = 4
                 .Col = 9
             Else
                 .TextMatrix(ActRow, 9) = ""
                 .TextMatrix(ActRow, 0) = ""
                 .Col = 9
             End If
        ElseIf .Col = 8 Then
            .TextMatrix(ActRow, 8) = txtOrder.Text
             txtOrder.Visible = False
             Call ComputeTotal
             If Val(.TextMatrix(ActRow, 8)) > 0 Then
                 .Row = ActRow
                 .Col = 0
                 '.CellFontName = "WINGDINGS"
                 '.CellFontSize = 12
                 '.Text = lblChk.Caption
                 '.CellAlignment = 4
                 .Col = 8
             Else
                 .TextMatrix(ActRow, 8) = ""
                 .TextMatrix(ActRow, 0) = ""
                 .Col = 8
             End If
        
        End If
        
     End With
End Sub
Sub ComputeTotal()
Dim nrw As Integer, GrnTl As Double, SPTotal As Double
    With flxSP
        SPTotal = Val(flxSP.TextMatrix(ActRow, 9)) * Val(Replace(flxSP.TextMatrix(ActRow, 10), ",", ""))
        flxSP.TextMatrix(ActRow, 11) = Format(SPTotal, "###,###,##0.#0")
        For nrw = 1 To .Rows - 1
            GrnTl = GrnTl + Val(Replace(flxSP.TextMatrix(nrw, 11), ",", ""))
        Next nrw
    End With
    txtTotalAmount.Text = Format(GrnTl, "###,###,##0.#0")
End Sub
Sub ComputeTotalView()
Dim nrw As Integer, GrnTl As Double, SPTotal As Double
    
    With flxSP
        For nrw = 1 To .Rows - 1
            SPTotal = Val(flxSP.TextMatrix(nrw, 9)) * Val(Replace(flxSP.TextMatrix(nrw, 10), ",", ""))
            flxSP.TextMatrix(nrw, 11) = Format(SPTotal, "###,###,##0.#0")
            GrnTl = GrnTl + Val(Replace(flxSP.TextMatrix(nrw, 11), ",", ""))
        Next nrw
    End With
    txtTotalAmount.Text = Format(GrnTl, "###,###,##0.#0")
End Sub
