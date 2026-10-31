VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form TRA 
   Caption         =   "Trust Receipt Agreement"
   ClientHeight    =   8805
   ClientLeft      =   510
   ClientTop       =   -7215
   ClientWidth     =   13335
   ControlBox      =   0   'False
   Icon            =   "TrustReceipt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   13335
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdOR 
      Height          =   375
      Left            =   9240
      TabIndex        =   85
      Tag             =   "XYF"
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Allocate O.R."
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
   Begin VB.CheckBox chkClub 
      Caption         =   "Yamaha Club member"
      Height          =   255
      Left            =   4320
      TabIndex        =   83
      Tag             =   "XYF"
      Top             =   7560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtSPinvID 
      Height          =   285
      Left            =   11760
      TabIndex        =   79
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFrontliner 
      Height          =   285
      Left            =   1560
      TabIndex        =   78
      Tag             =   "XYF"
      Top             =   7560
      Width           =   2415
   End
   Begin VB.TextBox txtRemarks 
      Height          =   645
      Left            =   1560
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   77
      Tag             =   "XYF"
      Top             =   7965
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   57
      Tag             =   "XYF"
      Top             =   720
      Width           =   13095
      Begin VB.ComboBox cboTrnsType 
         Height          =   315
         Left            =   1320
         TabIndex        =   80
         Tag             =   "XYF"
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboCustomer 
         Height          =   315
         Left            =   1320
         TabIndex        =   74
         Tag             =   "XYF"
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txtTerm 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   11400
         TabIndex        =   72
         Tag             =   "XYF"
         Top             =   960
         Width           =   1575
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
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   60
         Tag             =   "XYF"
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtInvNo 
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
         Left            =   11400
         MaxLength       =   10
         TabIndex        =   59
         Tag             =   "XYF"
         Top             =   248
         Width           =   1575
      End
      Begin VB.TextBox txtInvDate 
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
         Left            =   11400
         TabIndex        =   58
         Tag             =   "XYF"
         Top             =   600
         Width           =   1575
      End
      Begin lvButton.lvButtons_H cmdPrnFin 
         Height          =   375
         Left            =   5520
         TabIndex        =   87
         Tag             =   "xyf"
         Top             =   210
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   661
         Caption         =   "&Print Final"
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
         cFore           =   255
         cFHover         =   255
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "TrustReceipt.frx":08CA
         cBack           =   16777215
      End
      Begin VB.Label lblOSBal 
         Caption         =   "Outstanding Balance:"
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
         Left            =   3840
         TabIndex        =   86
         Tag             =   "XYF"
         Top             =   1020
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label lblORno 
         Alignment       =   2  'Center
         DataMember      =   "lblORno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5520
         TabIndex        =   84
         Tag             =   "xyf"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   9840
         TabIndex        =   81
         Tag             =   "XYF"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "TRANS. TYPE:"
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
         Left            =   240
         TabIndex        =   73
         Tag             =   "XYF"
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "SOLD TO"
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
         TabIndex        =   64
         Tag             =   "XYF"
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "ADDRESS"
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
         Left            =   240
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "T.R.A.#"
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
         Left            =   9840
         TabIndex        =   62
         Tag             =   "XYF"
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE:"
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
         Left            =   9840
         TabIndex        =   61
         Tag             =   "XYF"
         Top             =   660
         Width           =   1335
      End
   End
   Begin VB.TextBox txtVat 
      Height          =   285
      Left            =   3000
      TabIndex        =   56
      Top             =   9480
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPec 
      Height          =   615
      Left            =   2040
      TabIndex        =   55
      Top             =   9840
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
      Left            =   1320
      TabIndex        =   54
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin TabDlg.SSTab tabWS 
      Height          =   4455
      Left            =   6480
      TabIndex        =   47
      Top             =   9120
      Visible         =   0   'False
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "TrustReceipt.frx":0E64
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
         TabIndex        =   53
         Top             =   3480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptCash 
         Caption         =   "CASH"
         Height          =   255
         Left            =   2040
         TabIndex        =   52
         Top             =   3480
         Width           =   975
      End
      Begin VB.ComboBox cboWS 
         Height          =   2910
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   48
         Top             =   480
         Width           =   3615
      End
      Begin lvButton.lvButtons_H cmdWSOK 
         Height          =   375
         Left            =   240
         TabIndex        =   49
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
         Image           =   "TrustReceipt.frx":0E80
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdWSCancel 
         Height          =   375
         Left            =   2040
         TabIndex        =   50
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
         Image           =   "TrustReceipt.frx":141A
         cBack           =   -2147483633
      End
      Begin VB.Label Label6 
         Caption         =   "Select Wholesale:"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxWDisc 
      Height          =   735
      Left            =   8760
      TabIndex        =   46
      Top             =   9600
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
      TabIndex        =   45
      Top             =   9840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12600
      Top             =   120
   End
   Begin VB.TextBox txtCPOid 
      Height          =   285
      Left            =   5280
      TabIndex        =   44
      Top             =   9240
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
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   9120
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CheckBox chkDP 
      Caption         =   "DP"
      Height          =   255
      Left            =   360
      TabIndex        =   41
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstDelete 
      Height          =   645
      ItemData        =   "TrustReceipt.frx":19B4
      Left            =   9000
      List            =   "TrustReceipt.frx":19B6
      TabIndex        =   40
      Top             =   9480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   2040
      TabIndex        =   39
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmpID 
      Height          =   285
      Left            =   11160
      TabIndex        =   38
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtOSBal 
      Height          =   285
      Left            =   6480
      TabIndex        =   37
      Top             =   10440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPONO 
      Height          =   285
      Left            =   7440
      TabIndex        =   36
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtBillAmt 
      Height          =   285
      Left            =   4200
      TabIndex        =   35
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayStart 
      Height          =   285
      Left            =   3360
      TabIndex        =   34
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtNPays 
      Height          =   285
      Left            =   2520
      TabIndex        =   33
      Top             =   10440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayScheme 
      Height          =   285
      Left            =   1560
      TabIndex        =   32
      Top             =   10440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTerms 
      Height          =   285
      Left            =   480
      TabIndex        =   31
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   600
      TabIndex        =   30
      Top             =   9480
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
      Left            =   8880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   9120
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txtNoDiscount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "No Discount"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtGrandTotal 
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
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   8055
      Width           =   1620
   End
   Begin VB.TextBox txtLessDiscAmt 
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
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   7770
      Width           =   1620
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
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   7485
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   22
      Tag             =   "XYF"
      Text            =   "0%"
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox txtTtlDisc 
      Height          =   285
      Left            =   10080
      TabIndex        =   20
      Top             =   8880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9000
      TabIndex        =   19
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   8280
      TabIndex        =   18
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   315
      Left            =   11520
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      Image           =   "TrustReceipt.frx":19B8
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtSPSorderID 
      Height          =   285
      Left            =   720
      TabIndex        =   15
      Top             =   8880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   5160
      TabIndex        =   14
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   9120
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   9360
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   9120
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   9360
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   8880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1815
      Left            =   4680
      TabIndex        =   12
      Top             =   9120
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
      Left            =   6240
      TabIndex        =   11
      Top             =   9240
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
   Begin VB.TextBox txtBalance 
      Height          =   285
      Left            =   6480
      TabIndex        =   10
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Tag             =   "XY"
      Top             =   10920
      Width           =   135
   End
   Begin VB.Timer tmrQtyTip 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   9240
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Text            =   "Qty. On Hand: 0"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFlx 
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   9240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   2520
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   $"TrustReceipt.frx":1F52
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
            Picture         =   "TrustReceipt.frx":1FE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TrustReceipt.frx":2EBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TrustReceipt.frx":3D97
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TrustReceipt.frx":4C71
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TrustReceipt.frx":5B4B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxBO 
      Height          =   1095
      Left            =   7560
      TabIndex        =   43
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
      Index           =   0
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "TrustReceipt.frx":6A25
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   66
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "TrustReceipt.frx":6FBF
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   67
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Order"
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
      Image           =   "TrustReceipt.frx":7559
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   68
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Form"
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
      Image           =   "TrustReceipt.frx":7AF3
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "TrustReceipt.frx":808D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   70
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "TrustReceipt.frx":8627
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   71
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "TrustReceipt.frx":8BC1
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPrn 
      Height          =   1095
      Left            =   0
      TabIndex        =   82
      Top             =   9000
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   7
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "AGENT:"
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
      Left            =   360
      TabIndex        =   76
      Tag             =   "XYF"
      Top             =   7575
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Index           =   5
      Left            =   360
      TabIndex        =   75
      Tag             =   "XYF"
      Top             =   7980
      Width           =   1095
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
      Left            =   9510
      TabIndex        =   24
      Tag             =   "XYF"
      Top             =   8085
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
      Left            =   9840
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   7800
      Width           =   900
   End
   Begin VB.Label lblLess 
      AutoSize        =   -1  'True
      Caption         =   "Less:  "
      Height          =   195
      Left            =   8760
      TabIndex        =   21
      Tag             =   "XYF"
      Top             =   7815
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
      TabIndex        =   17
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   11040
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13575
      m_MDI_Height    =   9375
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
      Left            =   10080
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   7560
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
      TabIndex        =   3
      Top             =   2385
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE VAN"
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
      TabIndex        =   1
      Top             =   450
      Width           =   4000
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE VAN"
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
      TabIndex        =   2
      Top             =   480
      Width           =   4000
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
      TabIndex        =   4
      Top             =   2415
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
Attribute VB_Name = "TRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsCustomer As New ADODB.Recordset

Dim PrtSrch As Boolean, ActSrch As Boolean, PrtBln As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim tmrCount As Integer, BtnIndx As Integer
Dim AddNew As Boolean, PmanSrch As Boolean
Public JustScroll As Boolean
Dim fRow As Integer, fCol As Integer
Public EMPID As Integer
Dim wCost           As Variant

Private Sub BtnSaveCancel(bln As Boolean)
    Dim i As Integer
    
    If bln Then
        For i = 0 To 4
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(5).Left = 0
        cmdTrans(5).Visible = True
        cmdTrans(6).Left = cmdTrans(4).Width
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    End If
End Sub

Private Sub ClearEntries()
    cboCustomer.Text = ""
    txtAddress.Text = ""
    cboTrnsType.Text = ""
    txtInvNo.Text = ""
    txtInvDate.Text = ""
    txtRemarks.Text = ""
    txtTotal.Text = ""
    txtCustID.Text = ""
    txtTtlDisc.Text = ""
    txtTerms.Text = ""
    txtPayScheme.Text = ""
    txtNPays.Text = ""
    txtPayStart.Text = ""
    txtBillAmt.Text = ""
    txtOSBal.Text = ""
    txtGrandTotal.Text = ""
    txtLessDiscAmt.Text = ""
    flxOrd.Rows = 1
    flxOrd.Rows = 2
    txtPrintCopy = ""
'    txtCusType.Text = ""
    lblORno.Caption = ""
    cmdOR.Visible = False
End Sub
                                                                
Private Sub LockData(bln As Boolean)
'    Frame1.Enabled = Not Bln
    cboCustomer.Locked = bln
    txtAddress.Locked = bln
    cboTrnsType.Locked = bln
    txtInvNo.Locked = bln
    txtInvDate.Locked = bln
    txtTerm.Locked = bln
    txtRemarks.Locked = bln
    txtFrontliner.Locked = bln
    txtPctDisc.Locked = bln
    txtLessDiscAmt.Locked = bln
    JustScroll = bln
End Sub
   
Private Sub cboCustomer_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
        
    If PrtSrch Then
        PrtSrch = False
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
            txtAddress.Text = IIf(IsNull(rsCustomer!Address), "", rsCustomer!Address) '& IIf(IsNull(rsCustomer!CITYTOWN), "", " " & rsCustomer!CITYTOWN) & IIf(IsNull(rsCustomer!PROVINCE), "", " " & rsCustomer!Address)
            txtOSBal.Text = Format$(rsCustomer!OSbal, "#,###,##0.00")
            txtCusType.Text = IIf(IsNull(rsCustomer!Ctype), "", rsCustomer!Ctype)
            txtTerm.Text = IIf(IsNull(rsCustomer!Terms), "", rsCustomer!Terms)
            If Val(txtTerm.Text) = 0 Then
                txtTerm.Text = 60
            End If
            'If InStr(txtCusType.Text, "WHOLESALE") > 0 Then
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PrtSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtAddress.Text = ""
            txtOSBal.Text = ""
            txtCusType.Text = ""
            txtCustID.Text = ""
            txtTerm.Text = ""
        End If
    End With

End Sub

Private Sub cboCustomer_Click()
    Call cboCustomer_Change
End Sub


Private Sub cboCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PrtSrch = True
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAddress.SetFocus
End Sub

Private Sub cboTrnsType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtInvNo.SetFocus
End Sub

Private Sub cboWS_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
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


Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub


Private Sub GetCustomerDefaultTerm()
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, Gtotal As Currency
    
    'Saving On Account Job Order to On Account Terms
    'Get Terms
    '***Access
    'Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
    'mySQL
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

Private Sub cmdFind_Click()
    SPfind.MySource = "TRA"
    SPfind.Show vbModal
End Sub

Private Sub cmdOR_Click()
    Dim rsORsp As New ADODB.Recordset
    Dim MyOR As String, Qry As String, CustName As String, ValMsg As String, ORdate As String, myRem As String
    Dim ORAmt As Currency, ORalloc As Currency, AmtAlloc As Currency, SPttl As Currency, OrderID As Long, Ans As Integer
    
    MyOR = InputBox("Please enter your Official Receipt [O.R.] number:", "O.R. of Deposit for Spare Parts")
    If MyOR = "" Then Exit Sub
    
    ValMsg = ""
    CustName = ""
    ORAmt = 0
    ORalloc = 0
    OrderID = 0
    myRem = ""
    Qry = "SELECT * From payments Where payments.ORNO = '" & MyOR & "' AND payments.PayCode = 30;"
    With rsORsp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CustName = !CustomerName
            ORAmt = !Others
            ORalloc = IIf(IsNull(!SubAcctID), 0, !SubAcctID)
            OrderID = !OrderID
            ORdate = Format(!PayDate, "mm/dd/yyyy")
            myRem = IIf(IsNull(!Remarks), "", !Remarks)
            If ORAmt > 0 And OrderID > 0 And ORAmt = ORalloc Then
                Qry = "SELECT SPInvoice.SPINVNO From SPInvoice Where SPInvoice.SPINVOICEID = " & OrderID & ";"
                ValMsg = "O.R.# " & MyOR & " already allocated to Sales Order# " & GetFieldValue(Qry)
            End If
        Else
            ValMsg = "O.R.# " & MyOR & " not found!"
        End If
        .Close
    End With
    
    If ValMsg <> "" Then
        MsgBox ValMsg, vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Update Sales Order# " & txtInvNo.Text & " with the ff. details:" & vbCrLf & vbCrLf & _
            "O.R. Number: " & MyOR & "     O.R. Date: " & ORdate & vbCrLf & "O.R. Amount: " & Format(ORAmt - ORalloc, "#,##0.00") & vbCrLf & _
            "Customer Name: " & CustName, vbYesNo + vbQuestion, "Allocate O.R.")
    If Ans = vbNo Then Exit Sub
    
    SPttl = Val(Replace(txtGrandTotal.Text, ",", ""))
    AmtAlloc = ORAmt - ORalloc
    If AmtAlloc > SPttl Then AmtAlloc = SPttl
    If myRem <> "" Then
        myRem = myRem & ", " & txtInvNo.Text
    Else
        myRem = "Pd to SP Order#" & txtInvNo.Text
    End If
    LYRcon.Execute "UPDATE payments SET ORDERID = " & Val(txtSPinvID.Text) & ", SUBACCTID = " & ORalloc + AmtAlloc & ", REMARKS = '" & myRem & "' Where payments.ORNO = '" & MyOR & "';"
    If lblORno.Caption <> "" Then
        MyOR = Replace(lblORno.Caption, "O.R.#", "") & ", " & MyOR
    End If
    LYRcon.Execute "UPDATE SPInvoice SET ORNO = '" & MyOR & "', Balance = " & SPttl - AmtAlloc & " Where SPInvoice.SPINVOICEID = " & Val(txtSPinvID.Text) & ""
    
    lblORno.Caption = "O.R.# " & MyOR
    cmdOR.Visible = False
    MsgBox "O.R. successfully allocated!", vbInformation + vbOKOnly, "Sales Order"
End Sub

Private Sub cmdPrnFin_Click()
    Dim PrnVal(0 To 6) As String
        
    If txtInvNo.Text = "" Then
        MsgBox "No Available T.R.A. to Print", vbExclamation + vbOKOnly, "Print"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Print T.R.A.?", vbQuestion + vbYesNo, "Print FINAL")
    If Ans = vbNo Then Exit Sub
            
    PrnVal(0) = cboCustomer.Text
    PrnVal(1) = txtAddress.Text
    PrnVal(2) = txtTotal.Text
    dD = Day(DateValue(txtInvDate.Text))
    If Right(str(dD), 1) = "1" Then
        PrnVal(3) = dD & "st Day"
    ElseIf Right(str(dD), 1) = "2" Then
        PrnVal(3) = dD & "nd Day"
    ElseIf Right(str(dD), 1) = "3" Then
        PrnVal(3) = dD & "rd Day"
    Else
        PrnVal(3) = dD & "th Day"
    End If
    PrnVal(4) = MonthName(Month(DateValue(txtInvDate.Text)))
    PrnVal(5) = Right$(txtInvDate.Text, 2)
    PrnVal(6) = Format(Now, "mm/dd/yyyy hh:ss ampm")
    flxPrn.Rows = 1
    With flxOrd
        For i = 1 To .Rows - 1
            flxPrn.AddItem vbTab & .TextMatrix(i, 2) & vbTab & .TextMatrix(i, 3) & vbTab & .TextMatrix(i, 4) & "  " & .TextMatrix(i, 5) & vbTab & .TextMatrix(i, 7) & vbTab & .TextMatrix(i, 8)
        Next i
    End With

    Call PrintMyForm(PrnVal, "Trust Receipt Agreement", flxPrn)
    If Val(txtPrintCopy.Text) = 0 Then
        Call UpdateInvoice
    Else
        txtPrintCopy.Text = Val(txtPrintCopy.Text) + 1
        LYRcon.Execute "UPDATE spinvoice SET spinvoice.PRINTCOPY = " & Val(txtPrintCopy.Text) & " WHERE SPSorderID=" & Val(txtSPSorderID.Text) & ";"
    End If
    
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer, Qry As String, MyAcc As String, dD As Integer
    Dim myMsg As String
    
    'On Error GoTo LocalErr
    Select Case Index
        Case 0  'New
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=79);"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            If rsParts.State = adStateOpen Then rsParts.Close
            Timer1.Enabled = True
            Call ClearEntries
            Call LockData(False)
            AddNew = True
            cmdPrnFin.Visible = False
            Call BtnSaveCancel(True)
            cboTrnsType.ListIndex = 1
            flxOrd.SetFocus
            flxOrd.Col = 4
            txtInvDate.Text = Format(Date, "mm/dd/yyyy")
            cboCustomer.SetFocus
        Case 1  'Edit
            If txtInvNo.Text = "" Then
                MsgBox "No available Spareparts Invoice to edit", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=79);"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            
            If Val(txtPrintCopy.Text) > 0 Then
                MsgBox "Unable to edit Invoice# " & txtInvNo.Text & " because it is already printed.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If InStr(txtRemarks.Text, "CANCELLED") > 0 Then
                MsgBox "Unable to edit Invoice# " & txtInvNo.Text & " because it is already Cancelled.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            cmdTrans(1).Tag = txtInvNo.Text
            cmdPrnFin.Visible = False
            If rsParts.State = adStateOpen Then rsParts.Close
            Timer1.Enabled = True
            Call BtnSaveCancel(True)
            AddNew = False
            Call LockData(False)
        Case 2  'Summary
            FrmSummary.mySrc = "Trust Receipt Agreement"
            FrmSummary.Show
            Unload Me
        Case 3  'Print
            If txtInvNo.Text = "" Then
                MsgBox "No available TRA Sales Order to print", vbOKOnly + vbExclamation, "Print"
                Exit Sub
            End If
            If Val(txtPrintCopy.Text) > 0 Then
                Qry = "SELECT spitr.OUT, spitr.UNITM, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spitr.PRICE, spitr.PRICE * spitr.OUT AS AMOUNT " & _
                      "From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                      "Where spitr.REFERENCE = '" & txtInvNo.Text & "' AND spitr.TRANSTYPE = 'OUT' Order By spitr.`ORDER`"
            Else
                Qry = "SELECT spitr.OUT, spitr.UNITM, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spitr.PRICE, spitr.PRICE * spitr.OUT AS AMOUNT " & _
                      "From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
                      "Where spitr.REFERENCE = '" & txtInvNo.Text & "' AND spitr.TRANSTYPE = 'TEMPOUT' Order By spitr.`ORDER`"
            End If
            PRNatt.MySource = "TRA"
            PRNatt.QryPrn = Qry
            PRNatt.ParamValue1 = SysBranch
            PRNatt.ParamValue2 = cboCustomer.Text
            PRNatt.ParamValue3 = txtAddress.Text
            PRNatt.ParamValue4 = txtInvNo.Text
            PRNatt.ParamValue5 = txtInvDate.Text
            PRNatt.ParamValue6 = txtTerm.Text
            PRNatt.ParamValue7 = txtFrontliner.Text
            PRNatt.Show vbModal
        Case 4  'Exit
            'marketing.tbMain.Visible = True
            'ProperExit = True
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            'LogIn.Show
        Case 5  'Save
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            myMsg = ValidateSalesOrder
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save order?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Save Spare Parts Invoice
            
            Call SaveSPInvoice
            Call BtnSaveCancel(False)
            Call LockData(True)
            MsgBox "Spareparts Invoice Successfully Saved!", vbOKOnly + vbInformation, "Save"
            If Val(txtPrintCopy.Text) = 0 Then cmdPrnFin.Visible = True
        
        Case 6  'Cancel
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            Ans = MsgBox("Do you want to cancel Order?", vbQuestion + vbYesNo)
            If Ans = vbNo Then Exit Sub
            cmdPrnFin.Visible = False
            Call BtnSaveCancel(False)
            Call ClearEntries
            Call LockData(True)
    End Select
'    Exit Sub
'LocalErr:
'    MsgBox "Error Code: " & CStr(Err.Number) & " " & Err.Description, vbCritical + vbOKOnly, "Erro"
    
End Sub

Private Sub cmdTrans_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    'SPfind.MySource = "SPSORDER"
    'If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    'If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    'If tabWS.Tag <> "" Then Exit Sub
    'If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
End Sub


Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xRow As Integer, xBal As Integer, i As Integer
    
    'Clear FlxStock
    'mySQL
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
                    xBal = xBal - rsRec!Receive
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
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(StartDate, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(EndDate, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.SPAREPARTID=" & MyID & ";"
    
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
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(StartDate, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(EndDate, "yyyy-mm-dd") & "';"
    
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

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
    Dim Qry As String, Ans As Integer
    
    If cboWS.ListIndex = -1 Then
        MsgBox "Please select a customer from the list!", vbExclamation + vbOKOnly, "Wholesale"
        Exit Sub
    End If
    txtCustName.Text = cboWS.List(cboWS.ListIndex)
    txtCustID.Text = cboWS.ItemData(cboWS.ListIndex)
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
    '***Access
    'Qry = "SELECT SETTINGS.* From SETTINGS " & _
          "WHERE (((SETTINGS.DESCRIPTION)='Wholesale Discount')) ORDER BY SETTINGS.VALUE;"
    'mySQL
    Qry = "SELECT SETTINGS.* From SETTINGS " & _
          "WHERE SETTINGS.DESCRIPTION='Wholesale Discount' ORDER BY SETTINGS.VALUE;"
    
    rsDisc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsDisc.RecordCount > 0 Then
        rsDisc.MoveFirst
        flxWDisc.Rows = 0
        Do While Not rsDisc.EOF
            flxWDisc.AddItem rsDisc!Remarks & vbTab & rsDisc!Value
            rsDisc.MoveNext
        Loop
    Else
        flxWDisc.Rows = 0
    End If
    rsDisc.Close
    With rsDisc
        '***Access
        'Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='SP Disclusion')) ORDER BY SETTINGS.REMARKS;"
        'mySQL
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION='SP Disclusion' ORDER BY SETTINGS.REMARKS;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSPec.Rows = 0
            Do While Not .EOF
                flxSPec.AddItem !Remarks & vbTab & !Value
                .MoveNext
            Loop
        Else
            flxSPec.Rows = 0
        End If
        .Close
    End With
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
    If flxOrd.Rows = 2 And flxOrd.TextMatrix(1, 10) = "" Then
        Exit Sub
    Else
        Ans = MsgBox("Do you want the existing orders to avail the wholesale discount?", vbQuestion + vbYesNo, "Wholesale")
        If Ans = vbNo Then Exit Sub
        Call ComputeWholesaleDiscount
    End If
End Sub

Public Sub ComputeWholesaleDiscount()
    Dim PctDisc As Currency, DiscPrice As Currency, i As Integer, j As Integer, myVat As Single
    Dim myRem As String, Qry As String
    
    With flxOrd
        For j = 1 To .Rows - 1
            If .TextMatrix(j, 10) <> "" Then
                PctDisc = 0
                
                Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(j, 10)) & "));"
                PctDisc = Val(GetFieldValue(Qry))
              
                'old code
                'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
                '    If Val(.TextMatrix(j, 10)) = Val(flxSPec.TextMatrix(i, 0)) Then
                '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
                '        Exit For
                '    End If
                'Next i
                
                If PctDisc = 0 Then
                    rsParts.MoveFirst
                    rsParts.Find "[SPAREPARTID] = " + .TextMatrix(j, 10)
                    If Not rsParts.EOF Then
                        myRem = rsParts!Remarks
                    Else
                        myRem = ""
                    End If
                    If myRem <> "" Then
                        For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
                            If myRem = flxWDisc.TextMatrix(i, 0) Then
                                PctDisc = Val(flxWDisc.TextMatrix(i, 1))
                                Exit For
                            End If
                        Next i
                    End If
                    If PctDisc > 0 Then
                        myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                        DiscPrice = (Val(.TextMatrix(j, 14)) * myVat) * PctDisc
                        If InStr(myRem, "GENUINE") > 0 Then
                            If DiscPrice < 10 Then
                                DiscPrice = DiscPrice + 3
                            ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
                                DiscPrice = DiscPrice + 2
                            Else
                                If InStr(str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
                            End If
                            DiscPrice = Int(DiscPrice)
                        Else
                            If InStr(str(DiscPrice), ".") > 0 Then
                                DiscPrice = DiscPrice + 1
                                DiscPrice = Int(DiscPrice)
                            End If
                        End If
                        flxOrd.TextMatrix(j, 7) = Format$(DiscPrice, "###,##0.00") 'Price
                    Else
                        flxOrd.TextMatrix(j, 7) = Format$(Val(.TextMatrix(j, 13)), "###,##0.00") 'Orig SRP
                    End If
                Else    'Display Amount from Exclusion List
                    flxOrd.TextMatrix(j, 7) = Format$(PctDisc, "###,##0.00") 'Price
                End If
            End If
        Next j
    End With
    Call ComputeTotal
End Sub

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
        Case 6  'Discount
            With txtDiscount
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
                Else
                    txtNoDiscount.Top = flxOrd.CellTop + flxOrd.Top + 250
                    txtNoDiscount.Left = flxOrd.CellLeft + flxOrd.Left
                    txtNoDiscount.Text = "No Discount"
                    txtNoDiscount.Visible = True
                    .ZOrder
                End If
            End With
        Case 7         'Allow Administrators to Change the Price ... just temporary
            With txtPrice
                If flxOrd.TextMatrix(ActRow, 4) <> "" Then 'And UserType = "Administrator" Then
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
            '***Access
            'LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'FOUND' " & _
                           "WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & " AND SPSORDERDETAIL.SPAREPARTID=" & flxOrd.TextMatrix(ActRow, 10) & ";"
            
        Else
            eMsg = "Updating Order Status [Item not yet in Showroom]"
            flxOrd.TextMatrix(ActRow, 0) = ""
            'LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE (((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ") AND ((SPSORDERDETAIL.SPAREPARTID)=" & flxOrd.TextMatrix(ActRow, 10) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPSORDERDETAIL SET SPSORDERDETAIL.STATUS = 'BLANK' " & _
                           "WHERE SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & " AND SPSORDERDETAIL.SPAREPARTID=" & flxOrd.TextMatrix(ActRow, 10) & ";"
        End If
    End If
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub


Private Sub flxOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxOrd_Click
End Sub
                                                  
Private Sub flxOrd_KeyUp(KeyCode As Integer, Shift As Integer)
    'SPfind.MySource = "SPSORDER"
    'If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    'If KeyCode = vbKeyF11 Then InsertPO.Show vbModal
    'If tabWS.Tag <> "" Then Exit Sub
    'If KeyCode = vbKeyF10 And Not JustScroll Then Call cmdWholesale_Click
End Sub

Private Sub flxOrd_LeaveCell()
    txtNoDiscount.Visible = False
    txtAppModel.Visible = False
    txtSQtyBck.Visible = False
End Sub

Private Sub flxOrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
        If Val(.TextMatrix(.Row, 9)) = 0 Then
            flxOrd.TextMatrix(ActRow, 7) = "" 'Price
            flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
        Else
            'If UCase$(txtCusType.Text) = "WHOLESALE" Or UCase$(txtCusType.Text) = "LGC MARKETING" Then
                PctDisc = 0
                Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(.Row, 1)) & "));"
                PctDisc = Val(GetFieldValue(Qry))
                   
                'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
                '    If Val(.TextMatrix(.Row, 1)) = Val(flxSPec.TextMatrix(i, 0)) Then
                '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
                '        Exit For
                '    End If
                'Next i
                
                If PctDisc = 0 Then
                    For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
                        If .TextMatrix(.Row, 10) = flxWDisc.TextMatrix(i, 0) Then
                            PctDisc = Val(flxWDisc.TextMatrix(i, 1))
                            Exit For
                        End If
                    Next i
                    If PctDisc > 0 Then
                        myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                        DiscPrice = (Val(.TextMatrix(.Row, 9)) * myVat) * PctDisc
                        If InStr(.TextMatrix(.Row, 10), "GENUINE") > 0 Then
                            If DiscPrice < 10 Then
                                DiscPrice = DiscPrice + 3
                            ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
                                DiscPrice = DiscPrice + 2
                            Else
                                If InStr(str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
                            End If
                            DiscPrice = Int(DiscPrice)
                        Else
                            If InStr(str(DiscPrice), ".") > 0 Then
                                DiscPrice = DiscPrice + 1
                                DiscPrice = Int(DiscPrice)
                            End If
                        End If
                        flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Price
                    Else
                        flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                    End If
                Else    'Display Amount from Exclusion List
                    flxOrd.TextMatrix(ActRow, 7) = Format$(PctDisc, "###,##0.00") 'Price
                End If
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            'Else
            '    flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
            '    flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            'End If
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
        If Val(.TextMatrix(.Row, 9)) = 0 Then
            flxOrd.TextMatrix(ActRow, 7) = "" 'Price
            flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
        Else
            If UCase$(txtCusType.Text) = "WHOLESALE" Or UCase$(txtCusType.Text) = "LGC MARKETING" Then
                PctDisc = 0
                Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(.Row, 1)) & "));"
                PctDisc = Val(GetFieldValue(Qry))
                
                'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
                '    If Val(.TextMatrix(.Row, 1)) = Val(flxSPec.TextMatrix(i, 0)) Then
                '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
                '        Exit For
                '    End If
                'Next i
                
                If PctDisc = 0 Then
                    For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
                        If .TextMatrix(.Row, 10) = flxWDisc.TextMatrix(i, 0) Then
                            PctDisc = Val(flxWDisc.TextMatrix(i, 1))
                            Exit For
                        End If
                    Next i
                    If PctDisc > 0 Then
                        myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                        DiscPrice = (Val(.TextMatrix(.Row, 9)) * myVat) * PctDisc
                        If InStr(.TextMatrix(.Row, 10), "GENUINE") > 0 Then
                            If DiscPrice < 10 Then
                                DiscPrice = DiscPrice + 3
                            ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
                                DiscPrice = DiscPrice + 2
                            Else
                                If InStr(str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
                            End If
                            DiscPrice = Int(DiscPrice)
                        Else
                            If InStr(str(DiscPrice), ".") > 0 Then
                                DiscPrice = DiscPrice + 1
                                DiscPrice = Int(DiscPrice)
                            End If
                        End If
                        flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Price
                    Else
                        flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                    End If
                Else    'Display Amount from Exclusion List
                    flxOrd.TextMatrix(ActRow, 7) = Format$(PctDisc, "###,##0.00") 'Price
                End If
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            Else
                flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
            End If
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
    
    '***Access
    'Qry = "SELECT EMPLOYEES.FIRSTNAME, EMPLOYEES.EMPID From Employees " & _
          "WHERE (((EMPLOYEES.EMPID)=" & UserID & "));"
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

Function GetInvoiceNo() As String
    Dim rsInvNo As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    '***Access
    'Qry = "SELECT Max(Val([SPINVNO])) AS MaxNo FROM SPINVOICE;"
    'mySQl
    Qry = "SELECT Max(round(SPINVOICE.SPINVNO)) AS MaxNo FROM SPINVOICE;"
    
    With rsInvNo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, .Fields(0) + 1)
        Else
            MyNo = 0
        End If
        .Close
    End With
    GetInvoiceNo = Format(MyNo, "0000")
End Function



Private Sub DisplayPendingOrders()
    Dim rsPending As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
          "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((SPSALESORDER.EMPID)=" & EmpId & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#)) ORDER BY SPSALESORDER.ORDERNO;"
    'mySQL
    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.DESTINATION From SPSALESORDER " & _
          "WHERE (SPSALESORDER.INVOICENO) Is Null AND SPSALESORDER.EMPID=" & EMPID & " AND SPSALESORDER.ORDERDATE='" & txtDate.Text & "' ORDER BY SPSALESORDER.ORDERNO;"
    
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

Public Sub ShowRecord(MyID As Long)
    Dim rsSPinv As New ADODB.Recordset
    Dim rsDtls As New ADODB.Recordset
    Dim Qry As String
    
    With rsSPinv
        Qry = "SELECT * From SPInvoice Where spinvoice.SPSORDERID = " & MyID & ";"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtInvNo.Text = !SPINVNO
            txtInvDate.Text = Format(!InvDate, "mm/dd/yyyy")
            cboCustomer.Text = !chargeTo
            txtAddress.Text = !Address
            txtGrandTotal.Text = Format(!Amount, "#,##0.00")
            cboTrnsType.Text = !InvType
            txtRemarks.Text = !Remarks
            txtFrontliner.Text = !Frontliner
            'cboCusType.Text = !Ctype
            'txtSPinvID.Text = !SPINVOICEID
            txtPrintCopy.Text = !PRINTCOPY
            txtSPSorderID.Text = !SPSORDERID
            txtCusType.Text = IIf(IsNull(!Ctype), "", !Ctype)
            txtCustID.Text = !CUSTOMERID
            If Val(txtPrintCopy.Text) = 0 Then
                cmdPrnFin.Visible = True
            Else
                cmdPrnFin.Visible = False
            End If
        Else
            .Close
            Exit Sub
        End If
        .Close
        
        Qry = "SELECT spsalesorder.PCTDISC, spsalesorder.LESSDISCPRICE, spsalesorder.TERMS From spsalesorder " & _
            "Where spsalesorder.SPSORDERID = " & Val(txtSPSorderID.Text) & ";"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtTerm.Text = !Terms
            txtPctDisc.Text = !PctDisc & "%"
            txtLessDiscAmt.Text = Format(!LESSDISCPRICE, "#,##0.00")
        End If
        .Close
    End With

    
    
    If Val(txtPrintCopy.Text) > 0 Then
        Qry = "SELECT spitr.OUT, spitr.UNITM, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spitr.DISCOUNT, spitr.PRICE, spareparts.LOCATION, spitr.SPAREPARTID, spareparts.APPLICMODEL, spareparts.MAXDISCOUNT, spareparts.Cost " & _
              "From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) Where spitr.REFERENCE = '" & txtInvNo.Text & "' AND spitr.TRANSTYPE = 'OUT' Order By spitr.`ORDER`;"
    Else
        Qry = "SELECT spitr.OUT, spitr.UNITM, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spitr.DISCOUNT, spitr.PRICE, spareparts.LOCATION, spitr.SPAREPARTID, spareparts.APPLICMODEL, spareparts.MAXDISCOUNT, spareparts.Cost " & _
              "From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) Where spitr.REFERENCE = '" & txtInvNo.Text & "' AND spitr.TRANSTYPE = 'TEMPOUT' Order By spitr.`ORDER`;"
    End If
        With rsDtls
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxOrd.Rows = 1
                Do While Not .EOF
                    flxOrd.AddItem vbTab & vbTab & !OUT & vbTab & !UNITM & vbTab & !PartNumber & vbTab & !Description & vbTab & !DISCOUNT & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!OUT * !Price, "#,##0.00") & vbTab & vbTab & !SparepartID & vbTab & !ApplicModel & vbTab & !MaxDiscount & vbTab & !Cost
                    .MoveNext
                Loop
            End If
            .Close
        End With
                      
    'Else
    '    Qry = "SELECT spinvdtls.QTY, spinvdtls.UNIT, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spinvdtls.DISCOUNT, spinvdtls.PRICE, spareparts.LOCATION, spinvdtls.SPAREPARTID, spareparts.APPLICMODEL, spareparts.MAXDISCOUNT, spareparts.STOREPRICE, spinvdtls.COST " & _
    '          "From spinvdtls INNER JOIN spareparts ON (spinvdtls.SPAREPARTID = spareparts.SPAREPARTID) Where spinvdtls.SPINVOICEID = " & Val(txtSPinvID.Text) & " ORDER BY spinvdtls.SEQNO;"
    '    With rsDtls
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        If .RecordCount > 0 Then
    '            .MoveFirst
    '            flxOrd.Rows = 1
    '            Do While Not .EOF
    '                flxOrd.AddItem vbTab & vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & !Discount & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Qty * !Price, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & !MaxDiscount & vbTab & !STOREPRICE & vbTab & !Cost
    '                .MoveNext
    '            Loop
    '        End If
    '        .Close
    '    End With
    'End If
     
    Call ComputeTotal
            
    
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   ' SPfind.MySource = "SPSORDER"
   ' If KeyCode = vbKeyF12 Then SPfind.Show vbModal
End Sub

Private Sub Form_Load()
    Dim rsPartsMan As New ADODB.Recordset
    Dim Qry As String
    
    ActRow = 0
    ActCol = 0
    JustScroll = True
    txtInvDate.Text = Format(Date, "mm/dd/yyyy")
    Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='VAT';"
    txtVat.Text = GetFieldValue(Qry)
    txtFrontliner.Text = getFrontLiner
    Qry = "SELECT settings.REMARKS From settings Where settings.DESCRIPTION = 'SYSTEM BRANCH'"
    lblHeading(0).Caption = GetFieldValue(Qry)
    lblHeading(1).Caption = lblHeading(0).Caption
    
'    cboType.AddItem "CASH"
'    cboType.AddItem "ON-ACCOUNT"
'    cboType.AddItem "RECONDITION"
    
'    If InStr(lblHeading(0).Caption, "HONDA") > 0 Then
'        chkClub.Visible = False
'    Else
'        chkClub.Visible = True
'    End If
    
    cboTrnsType.AddItem "CASH"
    cboTrnsType.AddItem "CREDIT"
    MotoRep.sBar.Panels(6).Text = "Trust Receipt Agreement"
    flxOrd.ToolTipText = "v15.7.11 wCst"
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
        cmdCancel.Visible = False
        'cmdCashier.Visible = False
        cmdWareHouse.Visible = False
    Else
        cmdTrans(3).Visible = False
        cmdTrans(4).Visible = False
        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = True
        cmdTrans(2).Visible = True
        Call LockData(True)
        'cmdPO.Visible = False
        'cmdWholesale.Visible = False
        'cmdReq.Visible = False
        'cmdAdvFind.Visible = False
        'cmdCancel.Visible = False
        'cmdCashier.Visible = False
        'cmdWareHouse.Visible = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If ProperExit Then
    '    If rsParts.State = adStateOpen Then rsParts.Close
    '    ProperExit = False
    'ElseIf JustScroll Then
    '    Call cmdTrans_Click(2)
    '    ProperExit = False
    'Else
    '    Cancel = True
    'End If
    If rsParts.State = adStateOpen Then rsParts.Close
    If rsCustomer.State = adStateOpen Then rsCustomer.Close
    MotoRep.sBar.Panels(6).Text = ""
    
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

Private Sub SaveSPInvoice()
    Dim i As Integer, myTtl As Single, Pctd As Currency, AmtDisc As Currency
    Dim eMsg As String, Qry As String, myDisc As Currency, LstUpdt As String
                                  
    'Saving Table Spare Parts Invoice
    'On Error GoTo LocalErr
    myTtl = Val(Replace(txtGrandTotal.Text, ",", ""))
    Pctd = Val(Replace(txtPctDisc.Text, "%", ""))
    AmtDisc = Val(Replace(txtLessDiscAmt, ",", ""))
    LstUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
    Qry = "SELECT COUNT(spsalesorder.SPSORDERID) As FIELD_1 From spsalesorder Where spsalesorder.ORDERDATE = '" & Format(Date, "yyyy-mm-dd") & "'"
    txtSPSorderID.Tag = Val(GetFieldValue(Qry)) + 1
    If AddNew Then
        'save Sales Order first
        LYRcon.Execute "INSERT INTO spsalesorder(ORDERNO, ORDERDATE, FRONTLINER, REFERENCE, REMARKS, GROSSPRICE, PCTDISC, LESSDISCPRICE, TOTALPRICE, CUSTOMERID, CTYPE, CUSTOMER, TERMS, ADDRESS, WORKSTATION) " & _
            "VALUES(" & Val(txtSPSorderID.Tag) & ", '" & Format(Date, "yyyy-mm-dd") & "', '" & txtFrontliner.Text & "', 'TRA#:" & txtInvNo.Text & "', '" & txtRemarks.Text & "', " & Val(Replace(txtTotal.Text, ",", "")) & _
            ", " & Pctd & ", " & AmtDisc & ", " & myTtl & ", " & Val(txtCustID.Text) & ", '" & txtCusType.Text & "', '" & cboCustomer.Text & "', " & Val(txtTerm.Text) & ", '" & txtAddress.Text & "', '" & LstUpdt & "')"
        Qry = "SELECT spsalesorder.SPSORDERID From spsalesorder Where spsalesorder.ORDERDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND spsalesorder.WORKSTATION = '" & LstUpdt & "' AND " & _
                "spsalesorder.CUSTOMERID = " & Val(txtCustID.Text) & " AND spsalesorder.ORDERNO = " & Val(txtSPSorderID.Tag) & ";"
        txtSPSorderID.Text = GetFieldValue(Qry)
        eMsg = "Saving New Record for Sparepart Invoice"
        LYRcon.Execute "INSERT INTO SPInvoice (SPINVNO, Transno, InvDate, ChargeTo, InvType, Address, Amount, Frontliner, EmpId, SPSORDERID, Ctype, ReqID, Remarks, CUSTOMERID, UpdatedBy, LastUpdate) " & _
        "VALUES ('TRA#:" & txtInvNo.Text & "', '" & "TRA#:" & txtInvNo.Text & "', '" & Format(txtInvDate.Text, "yyyy-mm-dd") & "', '" & cboCustomer.Text & "', '" & cboTrnsType.Text & "', '" & Replace(txtAddress.Text, "'", "") & "', " & myTtl & ", '" & txtFrontliner.Text & "', " & Val(txtCustID.Text) & ", " & Val(txtSPSorderID.Text) & ", '" & txtCusType.Text & "', 0, '" & txtRemarks.Text & "', " & Val(txtCustID.Text) & ",  " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        'Qry = "SELECT SPInvoice.SPINVOICEID From SPInvoice Where spinvoice.SPINVNO = 'TRA#:" & txtInvNo.Text & "' AND spinvoice.LASTUPDATE = '" & LstUpdt & "';"
        'txtSPinvID.Text = GetFieldValue(Qry)
    Else
        LYRcon.Execute "UPDATE spsalesorder SET ORDERDATE = '" & Format(txtInvDate.Text, "yyyy-mm-dd") & "', FRONTLINER = '" & txtFrontliner.Text & "', REFERENCE = 'TRA#:" & txtInvNo.Text & "', REMARKS = '" & txtRemarks.Text & "', PCTDISC = " & Pctd & ", LESSDISCPRICE = " & AmtDisc & ", TOTALPRICE = " & myTtl & ", EMPID = " & UserID & ", " & _
        "CUSTOMERID = " & Val(txtCustID.Text) & ", CTYPE = '" & txtCusType.Text & "', TERMS = " & Val(txtTerm.Text) & ", WORKSTATION = '" & LstUpdt & "', CUSTOMER = '" & cboCustomer.Text & "', ADDRESS = '" & txtAddress.Text & "' " & _
        "Where spsalesorder.SPSORDERID = " & Val(txtSPSorderID.Text) & ";"
    
        eMsg = "Saving Edited Record for Sparepart Invoice"
        If Left$(txtInvNo.Text, 3) <> "TRA" Then txtInvNo.Text = "TRA#:" & txtInvNo.Text
        LYRcon.Execute "UPDATE spinvoice SET SPINVNO = '" & txtInvNo.Text & "', TRANSNO = '" & txtInvNo.Text & "', INVDATE = '" & Format(txtInvDate.Text, "yyyy-mm-dd") & "', CHARGETO = '" & cboCustomer.Text & "', INVTYPE = '" & cboTrnsType.Text & "', ADDRESS = '" & Replace(txtAddress.Text, "'", "") & "', AMOUNT = " & myTtl & ", FRONTLINER = '" & txtFrontliner.Text & "', EMPID = " & UserID & ", " & _
        "SPSORDERID = " & Val(txtSPSorderID.Text) & ", CTYPE = '" & txtCusType.Text & "', REMARKS = '" & txtRemarks.Text & "', CUSTOMERID = " & Val(txtCustID.Text) & ", UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
        "Where SPInvoice.SPINVNO = '" & cmdTrans(1).Tag & "' AND spinvoice.SPSORDERID = " & Val(txtSPSorderID.Text) & ";"
    End If
        
    
'    cboCustomer.RemoveItem cboCustomer.ListIndex
    
    
    'Saving Table Spare Parts Order Details
    If Left$(txtInvNo.Text, 3) <> "TRA" Then txtInvNo.Text = "TRA#:" & txtInvNo.Text
    If Not AddNew Then
        LYRcon.Execute "DELETE From spitr Where spitr.REFERENCE = '" & cmdTrans(1).Tag & "' AND spitr.TRANSTYPE = 'TEMPOUT'"
    End If
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
            If InStr(.TextMatrix(i, 6), "%") > 0 Then
                myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
            Else
                myDisc = Val(.TextMatrix(i, 6))
            End If
             
            wCost = GetFieldValue("SELECT spareparts.cost From spareparts Where spareparts.SPAREPARTID = " & Val(.TextMatrix(i, 10)) & ";")
            LYRcon.Execute "INSERT INTO SPITR (Reference, TRANSTYPE, Remarks, TransDate, SparepartID, `OUT`, Price, Discount, UnitM, Source, UpdatedBy, LastUpdate, `Order`, cost) " & _
            "VALUES ('" & txtInvNo.Text & "', 'TEMPOUT', '" & cboTrnsType.Text & "', '" & Format(txtInvDate.Text, "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & Val(Replace(.TextMatrix(i, 6), ",", "")) & "', '" & .TextMatrix(i, 3) & "', '" & cboCustomer.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & i & ", " & wCost & ");"
              'old
'            LYRcon.Execute "INSERT INTO SPITR (Reference, TRANSTYPE, Remarks, TransDate, SparepartID, `OUT`, Price, Discount, UnitM, Cost, Source, UpdatedBy, LastUpdate, `Order`) " & _
'            "VALUES ('" & txtInvNo.Text & "', 'TEMPOUT', '" & cboTrnsType.Text & "', '" & Format(txtInvDate.Text, "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", '" & Val(Replace(.TextMatrix(i, 6), ",", "")) & "', '" & .TextMatrix(i, 3) & "', '" & Val(Replace(.TextMatrix(i, 14), ",", "")) & "', '" & cboCustomer.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & i & ");"

'            ErrMsg = "[874] Error updating SP Qty on hand"
'            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 2)) & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                           "WHERE SPAREPARTS.SPAREPARTID=" & Val(.TextMatrix(i, 10)) & ";"
            
'            eMsg = "Line 1465: Adding New Sales Invoice Details @ row " & i
'            LYRcon.Execute "INSERT INTO spinvdtls(SPINVOICEID, SPAREPARTID, QTY, UNIT, PRICE, DISCOUNT, COST, SEQNO) " & _
'            "VALUES(" & Val(txtSPinvID.Text) & ", " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 2)) & ", '" & .TextMatrix(i, 3) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", " & myDisc & ", " & Val(.TextMatrix(i, 14)) & ", " & i & ");"
        Next i
    End With
    
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Public Sub UpdateInvoice()
    Dim i As Integer, myDisc As Currency, PayStart As Date
    Dim MoTerm As Integer
    
    txtPrintCopy.Text = Val(txtPrintCopy.Text) + 1
    LYRcon.Execute "UPDATE spinvoice SET spinvoice.PRINTCOPY = " & Val(txtPrintCopy.Text) & " WHERE SPSorderID=" & Val(txtSPSorderID.Text) & ";"
    
    If cboTrnsType.Text = "CREDIT" Then
        If Val(txtTerm.Text) = 30 Or Val(txtTerm.Text) = 60 Or Val(txtTerm.Text) = 90 Or Val(txtTerm.Text) = 120 Or Val(txtTerm.Text) = 150 Then
            MoTerm = Val(txtTerm.Text) / 30
            PayStart = DateAdd("M", MoTerm, DateValue(txtInvDate.Text))
        Else
            PayStart = DateAdd("d", Val(txtTerm.Text), DateValue(txtInvDate.Text))
        End If
        If Left$(txtInvNo.Text, 3) <> "TRA" Then txtInvNo.Text = "TRA#:" & txtInvNo.Text
        LYRcon.Execute "INSERT INTO ONACCTERMS (SPINVNO, Terms, PayScheme, Npays, PayStart, BillingAmt, NxtBill, AmtDue, Status, TtlAmtPaid, BALANCE, SPSORDERID, CUSTOMERID, UpdatedBy, LastUpdate, AcctAmount, TransDate) " & _
        "VALUES ('" & txtInvNo.Text & "', " & Val(txtTerm.Text) & ", 'MONTHLY', 1, '" & Format(PayStart, "yyyy-mm-dd") & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", '" & Format(PayStart, "yyyy-mm-dd") & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", 'OPEN', 0, " & Val(Replace(txtTotal.Text, ",", "")) & ", " & Val(txtSPSorderID.Text) & ", " & Val(txtCustID.Text) & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", '" & Format(txtInvDate.Text, "yyyy-mm-dd") & "');"
        
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL+" & Val(Replace(txtTotal.Text, ",", "")) & _
                       " WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
    Else
        If Left$(txtInvNo.Text, 3) <> "TRA" Then txtInvNo.Text = "TRA#:" & txtInvNo.Text
        LYRcon.Execute "INSERT INTO Payments (PayDate, InvoiceNO, SpareParts, CustomerName, CustomerID, Remarks, EmpId, UpdatedBy, LastUpdate) " & _
        "VALUES ('" & Format(txtInvDate.Text, "yyyy-mm-dd") & "', '" & txtInvNo.Text & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", '" & cboCustomer.Text & "', " & Val(txtCustID.Text) & ", 'CASH:" & txtCusType.Text & "', " & UserID & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
    End If
    'mySQL
        
    With flxOrd
        For i = 1 To .Rows - 1
            If InStr(.TextMatrix(i, 6), "%") > 0 Then
                myDisc = Val(Left$(.TextMatrix(i, 6), Len(.TextMatrix(i, 6)) - 1))
            Else
                myDisc = Val(.TextMatrix(i, 6))
            End If
            If Left$(txtInvNo.Text, 3) <> "TRA" Then txtInvNo.Text = "TRA#:" & txtInvNo.Text
            LYRcon.Execute "UPDATE spitr SET TRANSTYPE = 'OUT' Where spitr.REFERENCE = '" & txtInvNo.Text & "' AND spitr.TRANSTYPE = 'TEMPOUT' AND spitr.SPAREPARTID = " & Val(.TextMatrix(i, 10)) & ";"
        
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 2)) & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                           "WHERE SPAREPARTS.SPAREPARTID=" & Val(.TextMatrix(i, 10)) & ";"
        Next i
    End With
    'cmdOR.Visible = True
End Sub

Function GetOrderNo() As Integer
    Dim rsMyOrder As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
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

Function ValidateSalesOrder() As String
    Dim MyErr As String, Qry As String
    Dim i As Long, j As Long
    
    MyErr = ""
    If txtCustID.Text = "" Or cboCustomer.Text = "" Then MyErr = "Incomplete entry on field SOLD TO (CUSTOMER)"
    If txtAddress.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field ADDRESS"
    If txtInvNo.Text = "" Then
        MyErr = MyErr & vbCrLf & "Incomplete entry on field T.R.A. NO."
    ElseIf AddNew Or (Not AddNew And cmdTrans(1).Tag <> txtInvNo.Text) Then
        Qry = "SELECT SPInvoice.TRANSNO From SPInvoice Where spinvoice.SPINVNO = 'TRA#:" & txtInvNo.Text & "';"
        If CtrlNumExist(Qry) Then MyErr = MyErr & vbCrLf & "T.R.A. NO. already exist"
    End If
    If Not IsDate(txtInvDate.Text) Then MyErr = MyErr & vbCrLf & "Invalid entry on field TRA Date"
    If cboTrnsType.Text <> "CASH" And cboTrnsType.Text <> "CREDIT" Then MyErr = MyErr & vbCrLf & "Invalid TRANSACTION TYPE, please select from the list"
    If txtFrontliner.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field AGENT"
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
    
    ValidateSalesOrder = MyErr
End Function

Private Sub LoadSpareParts()
    Dim rsSet As New ADODB.Recordset
    Dim Qry As String
    Dim pPct As Single
        

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
End Sub


Private Sub RXResizer1_ResizeComplete()
    
    On Error GoTo LocalErr
    
    
    '1-Avl., 2-QTY., 3-UNIT, 4-PART NUMBER, 5-DESCRIPTION, 6-DISCOUNT, 7-UNIT PRICE, 8-TOTAL  PRICE, 9-LOCATION
    '10-SPAREPARTID, 11-APPLIC. MODEL, 12-maximum DISCOUNT, 13-ORIGINAL SRP, 14-COST

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
        .Col = 6: .Row = 1
        txtTotal.Left = flxOrd.Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7)
        txtTotal.Width = .ColWidth(8)
        flxParts.ColWidth(2) = .ColWidth(4)
        Label3.Left = txtTotal.Left - (Label3.Width + 50)
        Label3.Top = txtTotal.Top
        Label3.FontSize = txtTotal.FontSize
    End With
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

Private Sub Timer1_Timer()
    Dim Qry As String
    
    If sysBranchCode = "MWS" Then
'    - - - - - - - - - Reflect all customer name in TRA 10/24/2026- - - - - - - -
        Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL, CUSTOMER.CTYPE, CUSTOMER.TERMS, CUSTOMER.CITYTOWN, CUSTOMER.PROVINCE, CUSTOMER.CRLIMIT " & _
              "From CUSTOMER  ORDER BY CUSTOMER.CUSTOMERNAME;"
    Else
        Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL, CUSTOMER.CTYPE, CUSTOMER.TERMS, CUSTOMER.CITYTOWN, CUSTOMER.PROVINCE, CUSTOMER.CRLIMIT " & _
              "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
    End If
    With rsCustomer
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboCustomer.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
    If sysBranchCode = "MWS" Then
        Qry = "call sp_wholesale.sp_tra_CWV('" & sysBranchCode & "')"
    Else
        Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DLRPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST, SPAREPARTS.REMARKS, SPAREPARTS.WSMAXDISCOUNT " & _
              "From SPAREPARTS WHERE DESCRIPTION NOT LIKE '%INACTIVE%' ORDER BY SPAREPARTS.PARTNUMBER;"
    End If
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
    End With
    
'    Qry = "call sp_mrcretails.sp_TRA_getCustomer('" & sysBranchCode & "')"
'    With rsCustomer
'        If .State = adStateOpen Then .Close
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'                cboCustomer.AddItem !CustomerName
'                .MoveNext
'            Loop
'        End If
'    End With
'
'    Qry = "call sp_mrcretails.sp_TRA_getSpareparts('" & sysBranchCode & "')"
'    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
 '   Set flxParts.DataSource = rsParts
'    flxParts.Refresh
'    With flxParts
'        .Width = flxOrd.ColWidth(4)
'        .ColWidth(0) = 0
'        .ColWidth(1) = 0
'        .ColWidth(2) = flxOrd.ColWidth(4)
'        .ColWidth(3) = 0
 '       .ColWidth(4) = 0
 '       .ColWidth(5) = 0
 '       .ColWidth(6) = 0
 '       .ColWidth(7) = 0
 '       .ColWidth(8) = 0
 '       .ColWidth(9) = 0
 '       .ColWidth(10) = 0
 '   End With
    Timer1.Enabled = False
End Sub

Private Sub tmrQtyTip_Timer()
    Dim MyCost As Currency, CostVat As Currency
    
    If tmrCount = 0 Then
        With txtQtyTip
            .Top = flxOrd.CellTop + flxOrd.Top + 290
            .Left = flxOrd.Left + flxOrd.ColWidth(0) + flxOrd.ColWidth(1) + flxOrd.ColWidth(2) + flxOrd.ColWidth(3) + flxOrd.ColWidth(4) + flxOrd.ColWidth(5) + flxOrd.ColWidth(6)
            If Val(txtVat.Text) = 0 Then
                .Text = "No VAT available"
            Else
                MyCost = Val(flxOrd.TextMatrix(ActRow, 14))
                CostVat = MyCost * (1 + (Val(txtVat.Text) / 100))
                .Text = "Cost: " & Format(CostVat, "#,##0.00")
            End If
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

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then cboTrnsType.SetFocus
End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then txtAddress.SetFocus
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

Private Sub txtINVDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtInvNo.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        txtTerm.SetFocus
       ' flxOrd.Col = 4  'partnumber
    End If
End Sub

Private Sub txtINVDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789/-"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
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

Public Sub ComputeTotal()
    Dim i As Integer
    Dim TtlPrice As Currency, GrndTtl As Currency, DiscAmt As Currency, TtlDiscAmt As Currency
    Dim Pctd As Single, myPrice As Currency
    
    With flxOrd
        GrndTtl = 0
        TtlDiscAmt = 0
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
        Next i
        txtTtlDisc.Text = TtlDiscAmt
        txtTotal.Text = Format$(GrndTtl, "#,##0.00")
        txtGrandTotal.Text = Format$(GrndTtl - Val(Replace(txtLessDiscAmt.Text, ",", "")), "#,##0.00")
    End With
            
End Sub

Private Sub txtLessDiscAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtRemarks.SetFocus

End Sub

Private Sub txtLessDiscAmt_LostFocus()
    Dim Pctd As Double, DiscAmt As Currency, TtlAmt As Currency
    Dim AmtDsc As Single
    
    
    TtlAmt = Val(Replace(txtTotal.Text, ",", ""))
    DiscAmt = Val(Replace(txtLessDiscAmt.Text, ",", ""))
    If TtlAmt = 0 Then
        txtLessDiscAmt.Text = ""
        Exit Sub
    End If
    Pctd = (DiscAmt / TtlAmt) * 100
    txtPctDisc.Text = Format(Pctd, "##0.00") & "%"
    txtGrandTotal.Text = Format(TtlAmt - DiscAmt, "#,##0.00")
    AmtDsc = Val(Replace(txtLessDiscAmt.Text, ",", ""))
    If AmtDsc > Val(Replace(txtTotal.Text, ",", "")) Then
        MsgBox "DISCOUNT Amount must not be more than Total Amount!", vbCritical + vbOKOnly, "Warning"
        txtPctDisc.Text = ""
        txtLessDiscAmt.Text = ""
        txtGrandTotal.Text = txtTotal.Text
    End If

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
            If Val(.TextMatrix(.Row, 9)) = 0 Then
                flxOrd.TextMatrix(ActRow, 7) = "" 'Price
                flxOrd.TextMatrix(ActRow, 13) = "" 'SRP
            Else
                'If UCase$(txtCusType.Text) = "WHOLESALE" Or UCase$(txtCusType.Text) = "LGC MARKETING" Then
                    PctDisc = 0
                    Qry = "SELECT SPAREPARTS.DLRPRICE From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(.Row, 1)) & "));"
                    PctDisc = Val(GetFieldValue(Qry))
                   
                    'For i = 0 To flxSPec.Rows - 1 'Get Pct Discount base on SP Exclusion list
                    '    If Val(.TextMatrix(.Row, 1)) = Val(flxSPec.TextMatrix(i, 0)) Then
                    '        PctDisc = Val(flxSPec.TextMatrix(i, 1))
                    '        Exit For
                    '    End If
                    'Next i
                    If PctDisc = 0 Then
                        For i = 0 To flxWDisc.Rows - 1  'Get Pct Discount base on SP Type
                            If .TextMatrix(.Row, 10) = flxWDisc.TextMatrix(i, 0) Then
                                PctDisc = Val(flxWDisc.TextMatrix(i, 1))
                                Exit For
                            End If
                        Next i
                        If PctDisc > 0 Then
                            myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                            DiscPrice = (Val(.TextMatrix(.Row, 9)) * myVat) * PctDisc
                            If InStr(.TextMatrix(.Row, 10), "GENUINE") > 0 Then
                                If DiscPrice < 10 Then
                                    DiscPrice = DiscPrice + 3
                                ElseIf DiscPrice >= 10 And DiscPrice <= 100 Then
                                    DiscPrice = DiscPrice + 2
                                Else
                                    If InStr(str(DiscPrice), ".") > 0 Then DiscPrice = DiscPrice + 1
                                End If
                                DiscPrice = Int(DiscPrice)
                            Else
                                If InStr(str(DiscPrice), ".") > 0 Then
                                    DiscPrice = DiscPrice + 1
                                    DiscPrice = Int(DiscPrice)
                                End If
                            End If
                            flxOrd.TextMatrix(ActRow, 7) = Format$(DiscPrice, "###,##0.00") 'Price
                        Else
                            flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                        End If
                    Else    'Display Amount from Exclusion List
                        flxOrd.TextMatrix(ActRow, 7) = Format$(PctDisc, "###,##0.00") 'Price
                    End If
                    flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
                'Else
                 '   flxOrd.TextMatrix(ActRow, 7) = Format$(Val(.TextMatrix(.Row, 4)), "###,##0.00") 'Pricegud
                  '  flxOrd.TextMatrix(ActRow, 13) = Val(.TextMatrix(.Row, 4)) 'SRP
                'End If
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
    Exit Sub
LocalErr:
    If Err.Number = 30009 Then
        flxParts.Visible = False
        txtPartNo.Visible = False
        If rsParts.State = adStateOpen Then rsParts.Close
        Timer1.Enabled = True
        MsgBox "Program is reloading Partnumber" & vbCrLf & "Please wait for a few seconds...", vbInformation + vbOKOnly, "Reloading"
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
    Dim i As Long, j As Long
    
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

Private Sub txtInvNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtAddress.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        txtInvDate.SetFocus
    End If
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then txtInvDate.SetFocus
End Sub

Private Sub txtPctDisc_Change()
    Dim Pctd As Single, DiscAmt As Currency, TtlAmt As Currency
    
    Pctd = Val(Replace(txtPctDisc.Text, "%", ""))
    TtlAmt = Val(Replace(txtTotal.Text, ",", ""))
    DiscAmt = TtlAmt * (Pctd / 100)
    txtLessDiscAmt.Text = Format(DiscAmt, "#,##0.00")
    txtGrandTotal.Text = Format(TtlAmt - DiscAmt, "#,##0.00")
    
End Sub

Private Sub txtPctDisc_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtRemarks.SetFocus

End Sub

Private Sub txtPctDisc_LostFocus()
    Dim Pctd As Single
    
    Pctd = Val(Replace(txtPctDisc.Text, "%", ""))
    If Pctd > 100 Then
        MsgBox "Invalid DISCOUNT value!", vbCritical + vbOKOnly, "Warning"
        txtPctDisc.Text = ""
        txtLessDiscAmt.Text = ""
        txtGrandTotal.Text = txtTotal.Text
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
    End If
End Sub


Private Sub txtPrice_LostFocus()
    Dim AmtDsc As Currency, LessAmt As Currency, Pctd As Double, myVat As Currency
    
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
            If chkClub.Value = 0 Then
                myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
                AmtDsc = Val(.TextMatrix(ActRow, 14)) * myVat
                If Val(Replace(txtPrice.Text, ",", "")) < Format(AmtDsc, "###0.00") Then
                    MsgBox "Invalid Price entered!" & vbCrLf & "Please follow the correct Price", vbOKOnly + vbExclamation, "Invalid Price"
                    flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
                    flxOrd.TextMatrix(ActRow, 6) = ""
                    txtPrice.Visible = False
                    Call ComputeTotal
                    Exit Sub
                End If
            End If
        End If
        LessAmt = Val(.TextMatrix(ActRow, 13)) * (Val(.TextMatrix(ActRow, 12)) / 100)
        AmtDsc = Val(.TextMatrix(ActRow, 13)) - LessAmt
        If UserType <> "Administrator" Or chkClub.Value = 0 Then
            If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then
                MsgBox "Price must not be less than maximum discounted Price: " & Format(AmtDsc, "#,##0.00") & "!", vbCritical + vbOKOnly, "Validation"
                txtPrice.Tag = "XcdQty"
                txtPrice.SelStart = 0
                txtPrice.SelLength = Len(txtPrice.Text)
                txtPrice.SetFocus
                Exit Sub
            End If
        End If
        If Val(Replace(txtPrice.Text, ",", "")) < Val(.TextMatrix(ActRow, 13)) Then
            Pctd = (Val(.TextMatrix(ActRow, 13)) - Val(Replace(txtPrice.Text, ",", ""))) / Val(.TextMatrix(ActRow, 13))
            If Val(Replace(txtPrice.Text, ",", "")) > 0 Then .TextMatrix(ActRow, 6) = Format(Pctd, "#0.000%")
        End If
        If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then 'Price is less than allowable discount
            .TextMatrix(ActRow, 3) = IIf(InStr(.TextMatrix(ActRow, 3), "*") > 0, "", .TextMatrix(ActRow, 3) & "*")
        Else
            .TextMatrix(ActRow, 3) = Replace(.TextMatrix(ActRow, 3), "*", "")
        End If
    End With
    flxOrd.TextMatrix(ActRow, 7) = Format(txtPrice.Text, "#,##0.00")
    Call ComputeTotal
    txtPrice.Visible = False

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
        LYRcon.Execute "Update settings SET `VALUE` = '" & myPvOrder & "', REMARKS = '" & Format(myOrderDt, "yyyy-mm-dd-mm") & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Date, "yyyy-mm-dd hh:mm:ss") & "' " & _
            "Where settings.DESCRIPTION = 'SALES ORDER';"
End Sub




