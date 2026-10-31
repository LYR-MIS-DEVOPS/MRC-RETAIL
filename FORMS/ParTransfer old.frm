VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ParTrans 
   Caption         =   "Spare Parts Transfer"
   ClientHeight    =   8145
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14010
   ControlBox      =   0   'False
   Icon            =   "ParTransfer old.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   14010
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtIntransit 
      Height          =   375
      Left            =   1680
      TabIndex        =   82
      Top             =   9000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtTotPen 
      Height          =   315
      Left            =   30
      TabIndex        =   80
      Tag             =   "XYF"
      Top             =   7110
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame frmSup 
      Caption         =   "SUPPLIER NAME"
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
      Height          =   855
      Left            =   11040
      TabIndex        =   75
      Tag             =   "XYF"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblSupplier 
         Alignment       =   2  'Center
         Caption         =   "----"
         Height          =   495
         Left            =   120
         TabIndex        =   76
         Tag             =   "XYF"
         Top             =   240
         Width           =   2415
      End
   End
   Begin lvButton.lvButtons_H cmdCardex 
      Height          =   315
      Left            =   3630
      TabIndex        =   74
      Tag             =   "XYF"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "&Cardex"
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
      Image           =   "ParTransfer old.frx":17D2A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtcorr 
      Height          =   315
      Left            =   6210
      TabIndex        =   71
      Top             =   9660
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1080
      Left            =   120
      TabIndex        =   60
      Tag             =   "XYF"
      Top             =   6900
      Width           =   13755
      Begin VB.TextBox txtAprvBy 
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
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   540
         Width           =   5415
      End
      Begin VB.TextBox txtPrepBy 
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
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   540
         Width           =   3615
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
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   200
         Width           =   11895
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
         Index           =   2
         Left            =   6915
         TabIndex        =   68
         Tag             =   "XYF"
         Top             =   600
         Width           =   1155
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
         Index           =   1
         Left            =   465
         TabIndex        =   67
         Tag             =   "XYF"
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   180
         TabIndex        =   61
         Tag             =   "XYF"
         Top             =   285
         Width           =   1395
      End
   End
   Begin VB.TextBox txtVat 
      Height          =   285
      Left            =   2760
      TabIndex        =   55
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPec 
      Height          =   615
      Left            =   2040
      TabIndex        =   54
      Top             =   11040
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
      TabIndex        =   53
      Top             =   10800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxWDisc 
      Height          =   735
      Left            =   8760
      TabIndex        =   52
      Top             =   10800
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
      TabIndex        =   51
      Top             =   11040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8880
      Top             =   0
   End
   Begin VB.TextBox txtCPOid 
      Height          =   285
      Left            =   5280
      TabIndex        =   50
      Top             =   10440
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
      TabIndex        =   48
      Top             =   9480
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CheckBox chkDP 
      Caption         =   "DP"
      Height          =   255
      Left            =   360
      TabIndex        =   47
      Top             =   10680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstDelete 
      Height          =   645
      ItemData        =   "ParTransfer old.frx":182C4
      Left            =   8790
      List            =   "ParTransfer old.frx":182C6
      TabIndex        =   46
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   1830
      TabIndex        =   45
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmpID 
      Height          =   285
      Left            =   10920
      TabIndex        =   44
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtOSBal 
      Height          =   285
      Left            =   6480
      TabIndex        =   43
      Top             =   11640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPONO 
      Height          =   285
      Left            =   7200
      TabIndex        =   42
      Top             =   9240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtBillAmt 
      Height          =   285
      Left            =   4200
      TabIndex        =   41
      Top             =   11640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayStart 
      Height          =   285
      Left            =   3360
      TabIndex        =   40
      Top             =   11640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtNPays 
      Height          =   285
      Left            =   2520
      TabIndex        =   39
      Top             =   11640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPayScheme 
      Height          =   285
      Left            =   1560
      TabIndex        =   38
      Top             =   11640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTerms 
      Height          =   285
      Left            =   480
      TabIndex        =   37
      Top             =   11640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   360
      TabIndex        =   36
      Top             =   10080
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
      TabIndex        =   35
      Top             =   9720
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
      TabIndex        =   34
      Text            =   "No Discount"
      Top             =   9600
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
      Left            =   8550
      Locked          =   -1  'True
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtTtlDisc 
      Height          =   285
      Left            =   9840
      TabIndex        =   32
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPacking 
      Height          =   285
      Left            =   8760
      TabIndex        =   31
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtBranchID 
      Height          =   285
      Left            =   8040
      TabIndex        =   30
      Top             =   9240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtInvNo 
      Height          =   285
      Left            =   5760
      TabIndex        =   29
      Top             =   9240
      Visible         =   0   'False
      Width           =   375
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   10230
      TabIndex        =   28
      Tag             =   "XYF"
      Top             =   1920
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
      Image           =   "ParTransfer old.frx":182C8
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   11940
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   1980
      Width           =   1815
   End
   Begin VB.TextBox txtSPSorderID 
      Height          =   285
      Left            =   480
      TabIndex        =   26
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   4920
      TabIndex        =   25
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   9480
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   9960
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   9480
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   9960
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   1560
      TabIndex        =   24
      Top             =   9240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   6120
      TabIndex        =   23
      Top             =   10200
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
      TabIndex        =   22
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Tag             =   "XY"
      Top             =   12120
      Width           =   135
   End
   Begin VB.Timer tmrQtyTip 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   9600
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   600
      TabIndex        =   20
      Text            =   "Qty. On Hand: 0"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFlx 
      Height          =   315
      Left            =   3720
      TabIndex        =   19
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Top             =   10440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   4350
      Left            =   120
      TabIndex        =   5
      Tag             =   "XY"
      Top             =   2400
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   7673
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      Enabled         =   -1  'True
      AllowUserResizing=   1
      FormatString    =   $"ParTransfer old.frx":18862
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
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1380
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   400
      Width           =   13755
      Begin VB.TextBox txtSQSERVE 
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
         Left            =   6330
         Locked          =   -1  'True
         TabIndex        =   79
         Tag             =   "XYF"
         Top             =   960
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtDate 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   240
         Width           =   5175
      End
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
         Height          =   255
         Left            =   11640
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   990
         Width           =   2055
      End
      Begin VB.TextBox txtRefNo 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtPO 
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
         MaxLength       =   100
         TabIndex        =   62
         Tag             =   "XYF"
         Top             =   960
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   600
         Width           =   5175
      End
      Begin lvButton.lvButtons_H cmdLoad 
         Height          =   375
         Left            =   7080
         TabIndex        =   77
         Tag             =   "XYF"
         Top             =   960
         Visible         =   0   'False
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   661
         Caption         =   "&Load S.O #"
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
         Image           =   "ParTransfer old.frx":188FE
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "SERVED:"
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
         Left            =   5520
         TabIndex        =   78
         Tag             =   "XYF"
         Top             =   1020
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   9840
         TabIndex        =   66
         Tag             =   "XYF"
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. NUMBER:"
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
         TabIndex        =   65
         Tag             =   "XYF"
         Top             =   990
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ADDRESS:"
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
         Left            =   120
         TabIndex        =   64
         Tag             =   "XYF"
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TRANSFERRED TO:"
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
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. DATE"
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
         Left            =   9960
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   630
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   4800
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
            Picture         =   "ParTransfer old.frx":18E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ParTransfer old.frx":19D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ParTransfer old.frx":1AC4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ParTransfer old.frx":1BB26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ParTransfer old.frx":1CA00
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxBO 
      Height          =   1095
      Left            =   7560
      TabIndex        =   49
      Top             =   11520
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
      TabIndex        =   0
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
      Image           =   "ParTransfer old.frx":1D8DA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   56
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
      Image           =   "ParTransfer old.frx":1DE74
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   57
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Form"
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
      Image           =   "ParTransfer old.frx":1E40E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   58
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
      Image           =   "ParTransfer old.frx":1E9A8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   9
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
      Image           =   "ParTransfer old.frx":1EF42
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   10
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
      Image           =   "ParTransfer old.frx":1F4DC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   59
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
      Image           =   "ParTransfer old.frx":1FA76
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdAdvFind 
      Height          =   375
      Left            =   8550
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "XYF"
      ToolTipText     =   "[F12] Advance Search"
      Top             =   1920
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Adv. Find"
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
      Image           =   "ParTransfer old.frx":20010
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   165
      Left            =   0
      TabIndex        =   69
      Tag             =   "XYF"
      Top             =   7920
      Visible         =   0   'False
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxOld 
      Height          =   3585
      Left            =   11010
      TabIndex        =   70
      Tag             =   "XY"
      Top             =   3150
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      FormatString    =   "^  DATE     |<REF. NO.  |^ OLD COST |<  OLD PRICE"
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
   Begin lvButton.lvButtons_H cmdCheck 
      Height          =   315
      Left            =   5310
      TabIndex        =   72
      Tag             =   "XYF"
      ToolTipText     =   "Add Costing"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "Check Costing"
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
      Image           =   "ParTransfer old.frx":205AA
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCheckUp 
      Height          =   315
      Left            =   6990
      TabIndex        =   73
      Tag             =   "XYF"
      ToolTipText     =   "Update Costing"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Caption         =   "Update Corrt'n"
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
      Image           =   "ParTransfer old.frx":20B44
      Enabled         =   0   'False
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
      Left            =   30
      TabIndex        =   81
      Top             =   8850
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14130
      m_MDI_Height    =   8655
   End
   Begin VB.Label lblTotal 
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
      Left            =   7800
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   9000
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSFERRED SPARE PARTS"
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
      Left            =   120
      TabIndex        =   15
      Top             =   2025
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSFERRED SPARE PARTS"
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
      Left            =   30
      TabIndex        =   16
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
Attribute VB_Name = "ParTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset

Dim PrtSrch As Boolean, ActSrch As Boolean, PrtBln As Boolean, ProperExit As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim tmrCount As Integer, BtnIndx As Integer
Dim AddNew As Boolean, PmanSrch As Boolean
Public JustScroll As Boolean
Dim fRow As Integer, fCol As Integer
Public EmpId As Integer
Dim rsPhist As New ADODB.Recordset

Dim OrdRows         As String
Dim pOrder          As Long
Dim wCost           As Variant
Dim showcost        As Boolean

Private Sub ClearEntries()
    cboBranch.Text = ""
    txtAddress.Text = ""
    txtPONO.Text = ""
    txtRefNo.Text = "<Auto>"
    txtDate.Text = ""
    chkFinal.Value = 0
    txtRemarks.Text = ""
    txtPrepBy.Text = ""
    txtAprvBy.Text = ""
    flxOrd.Rows = 1
    flxOrd.Rows = 2
    txtPrintCopy.Text = ""
    txtTotal.Text = ""
    txtPO.Text = ""
End Sub
                                                                
Private Sub LockData(bln As Boolean)
    Frame1.Enabled = bln
    Frame2.Enabled = bln
    JustScroll = Not bln
End Sub
   
Private Sub cboBranch_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If PmanSrch Then
        PmanSrch = False
        Exit Sub
    End If
    With cboBranch
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            txtBranchID.Text = .ItemData(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PmanSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtBranchID.Text = ""
        End If
    End With

End Sub

Private Sub cboBranch_Click()
    Call cboBranch_Change
End Sub

Private Sub cboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAddress.SetFocus
End Sub

Private Sub cboBranch_LostFocus()
    Dim Qry As String
    
    If Val(txtBranchID.Text) = 0 Then Exit Sub
    'Qry = "SELECT BRANCH.ADDRESS From BRANCH WHERE (((BRANCH.BRANCHID)=" & Val(txtBranchID.Text) & "));"
    'mySQL
    
    Qry = "SELECT BRANCH.ADDRESS From BRANCH WHERE BRANCH.BRANCHID=" & Val(txtBranchID.Text) & ";"
    
    txtAddress.Text = GetFieldValue(Qry)
End Sub




Private Sub cmdAdvFind_Click()
    SPfind.MySource = "SPTRANSFER"
    SPfind.Show vbModal
    'MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Order"
End Sub

Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub


Private Sub cmdCardex_Click()
    SPCardex.Show
End Sub
Private Sub cmdCheck_Click()
    If Val(flxOrd.TextMatrix(1, 2)) > 0 Then
        flxOrd.ColWidth(8) = (flxOrd.Width - 300) * 0.1
        flxOrd.TextMatrix(0, 8) = "CORRECT 'N"
        flxOrd.ColWidth(5) = (flxOrd.Width - 300) * 0.3
        cmdCheckUp.Enabled = True
    End If
End Sub
Private Sub cmdCheckUp_Click()
    Dim i As Long, hCost As Double
    Dim Ans As Integer
    
    Ans = MsgBox("Save Cost correction?", vbYesNo + vbQuestion, "Correction")
    If Ans = vbNo Then Exit Sub
    
    'Update cost
    LYRcon.Execute "UPDATE STOCKTRANSFER SET TOTALCOST = " & Val(Replace(txtTotal.Text, ",", "")) & ", REMARKS = '" & txtRemarks.Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
      "WHERE (STOCKTRANSFER.TRANSFERID =" & Val(txtRefNo.Text) & ");"
    
    'Update details Cost b4 final
    With flxOrd
        For i = 1 To flxOrd.Rows - 1
            If Val(.TextMatrix(i, 8)) > 0 Then
                 LYRcon.Execute "UPDATE stocktransdtl SET COST = " & Val(Replace(.TextMatrix(i, 8), ",", "")) & " Where stocktransdtl.TRANSFERID = " & Val(txtRefNo.Text) & " AND stocktransdtl.SPAREPARTID = " & Val(.TextMatrix(i, 11)) & ";"
            End If
            If chkFinal.Value = 1 Then
                If Val(.TextMatrix(i, 8)) > 0 Then
                    LYRcon.Execute "UPDATE  SPITR SET PRICE =" & Val(Replace(.TextMatrix(i, 8), ",", "")) & ", UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' WHERE SPITR.REFERENCE = 'TR#" & Val(txtRefNo.Text) & "' AND SPITR.SPAREPARTID = " & Val(.TextMatrix(i, 11)) & ";"
                End If
            End If
            DoEvents
        Next i
        MsgBox ("Cost Successfully Updated"), vbInformation, "Cost Updated"
        cmdCheckUp.Enabled = False
    End With
End Sub
Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxOrd
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

Private Sub cmdLoad_Click()
    Call LoadBranchPO
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer, MyRef As Integer
    Dim myMsg As String, PrnVal(0 To 12) As String, Qry As String
    
     
    'On Error GoTo LocalErr
    
    Select Case Index
        Case 0  'New
            
            Call RXResizer1_ResizeComplete
            Call ClearEntries
            Call LockData(True)
            'txtDate.Text = Format(Date, "MM/DD/YYYY")
            txtDate.Text = Format(ServerTime, "MM/DD/YYYY")
            txtRefNo.Text = "<Auto>"
            txtAprvBy.Text = UCase(SysBranchOIC)
            txtPrepBy.Text = UCase(UserName)
            'cboBranch.Locked = False
            AddNew = True
            Call ButtonSaveCancel(True)
            If rsParts.State = adStateOpen Then rsParts.Close
            Call ReloadSpareParts
            cboBranch.SetFocus
            chkFinal.Enabled = True
            cmdCheck.Enabled = False
            cmdCheckUp.Enabled = False
        Case 1  'Edit
            If txtRefNo.Text = "" Then
                MsgBox "No available record to edit", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If InStr(txtRemarks.Text, "CANCELLED") > 0 Then
                MsgBox "Unable to edit Ref. No. " & txtRefNo.Text & " because it was already Cancelled.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If chkFinal.Value = 1 Then
                MsgBox "Unable to edit Ref. No. " & txtRefNo.Text & " because record is already final.", vbOKOnly + vbExclamation, "Edit"
                Exit Sub
            End If
            If rsParts.State = adStateOpen Then rsParts.Close
            Call ReloadSpareParts
            AddNew = False
            Call LockData(True)
            Call ButtonSaveCancel(True)
            chkFinal.Enabled = True
            cboBranch.Locked = False
        Case 2 ' Summary
            TranSummary.Show
            Unload Me
'            MotoRep.toolBarButton.Visible = False
'            MotoRep.imgBG.Visible = False
        Case 3  'Print
        
            Dim showCostUser As String
    
    
            showCostUser = GetFieldValue("SELECT pcuser.DESIGNATION From pcuser Where pcuser.EmpId ='" & UserID & "'")
            If UCase(showCostUser) = "SYSTEM ADMIN." Or showCostUser = "AUDITOR" Then  'Desig = "SYSTEM ADMIN." Or Desig = "AUDITOR"
                showcost = True
            Else
                showcost = False
            End If
            
            If txtRefNo.Text = "" Then
                MsgBox "No Available Stock Transfer to Print", vbExclamation + vbOKOnly, "Print"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Print Stock Transfer?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            
            MyRef = Val(txtRefNo.Text)
            
            If showcost And chkFinal.Value = 1 Then
                'PRNmc.QryPrn = "SELECT SPITR.OUT, IIf([spitr].[out]>1,'pcs.','pc.') AS unit, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [PRICE]*[OUT] AS TOTAL " & _
                               "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='TR#" & Val(txtRefNo.Text) & "') AND ((SPITR.REMARKS)='TRANSFER'));"
                'mySQL
                If txtIntransit.Text = "INTRANSIT" Then
'                    PRNmc.QryPrn = "SELECT SPITR.OUT, If(spitr.out>1,'pcs.','pc.') AS unit, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.OUT AS TOTAL " & _
'                               "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='TR#" & Val(txtRefNo.Text) & "' order by SPAREPARTS.PARTNUMBER;"

                    PRNmc.QryPrn = "SELECT SPITR.OUT, If(spitr.out>1,'pcs.','pc.') AS unit, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.COST, SPITR.COST*SPITR.OUT AS TOTAL " & _
                               "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='TR#" & Val(txtRefNo.Text) & "' order by SPAREPARTS.PARTNUMBER;"
                    
                Else
'                    PRNmc.QryPrn = "SELECT SPITR.OUT, If(spitr.out>1,'pcs.','pc.') AS unit, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.OUT AS TOTAL " & _
'                               "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='TR#" & Val(txtRefNo.Text) & "' AND SPITR.REMARKS='TRANSFER' order by SPAREPARTS.PARTNUMBER;"
                    
                    PRNmc.QryPrn = "SELECT SPITR.OUT, If(spitr.out>1,'pcs.','pc.') AS unit, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.COST, SPITR.COST*SPITR.OUT AS TOTAL " & _
                               "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='TR#" & Val(txtRefNo.Text) & "' AND SPITR.REMARKS='TRANSFER' order by SPAREPARTS.PARTNUMBER;"

                End If
                PRNmc.MySource = "STOCKTRANSFER W/ COST"
            ElseIf showcost And chkFinal.Value = 0 Then
                PRNmc.MySource = "STOCKTRANSFER W/ COST"
                'PRNmc.QryPrn = "SELECT STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, STOCKTRANSDTL.COST, [STOCKTRANSDTL.COST]*[QTY] AS TOTAL  " & _
                             "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((STOCKTRANSDTL.TRANSFERID) = " & MyRef & ")) ORDER BY STOCKTRANSDTL.SEQNO;"
                'mySQL
'                PRNmc.QryPrn = "SELECT STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, STOCKTRANSDTL.COST, STOCKTRANSDTL.COST*STOCKTRANSDTL.QTY AS TOTAL  " & _
'                             "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where STOCKTRANSDTL.TRANSFERID = " & MyRef & " ORDER BY STOCKTRANSDTL.SEQNO;"
                PRNmc.QryPrn = "SELECT STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, STOCKTRANSDTL.COST, STOCKTRANSDTL.COST*STOCKTRANSDTL.QTY AS TOTAL  " & _
                             "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where STOCKTRANSDTL.TRANSFERID = " & MyRef & " order by SPAREPARTS.PARTNUMBER;"

            
            Else
                PRNmc.MySource = "STOCKTRANSFER"
                'PRNmc.QryPrn = "SELECT STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, STOCKTRANSDTL.PACKING " & _
                             "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((STOCKTRANSDTL.TRANSFERID) = " & MyRef & ")) ORDER BY STOCKTRANSDTL.SEQNO;"
                'mySQL
                If chkFinal.Value = 1 Then
'                    PRNmc.QryPrn = "SELECT STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, STOCKTRANSDTL.PACKING " & _
'                                 "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where STOCKTRANSDTL.TRANSFERID = " & MyRef & " ORDER BY STOCKTRANSDTL.SEQNO;"
                      PRNmc.QryPrn = "SELECT STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, STOCKTRANSDTL.PACKING " & _
                                 "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where STOCKTRANSDTL.TRANSFERID = " & MyRef & " order by SPAREPARTS.PARTNUMBER;"
               Else
'                    PRNmc.QryPrn = "SELECT STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION " & _
'                                 "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where STOCKTRANSDTL.TRANSFERID = " & MyRef & " ORDER BY STOCKTRANSDTL.SEQNO;"
                    PRNmc.QryPrn = "SELECT STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION " & _
                                 "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where STOCKTRANSDTL.TRANSFERID = " & MyRef & " order by SPAREPARTS.PARTNUMBER;"
                End If
            End If
            'Qry = "SELECT BRANCH.BRANCHNAME From BRANCH WHERE (((BRANCH.SYSTEMBRANCH)=True));"
            'mySQL
            Qry = "SELECT BRANCH.BRANCHNAME From BRANCH WHERE BRANCH.SYSTEMBRANCH=True;"
            
            PRNmc.ParamValue1 = GetFieldValue(Qry)
            'Qry = "SELECT BRANCH.ADDRESS From BRANCH WHERE (((BRANCH.SYSTEMBRANCH)=True));"
            'mySQL
            Qry = "SELECT BRANCH.ADDRESS From BRANCH WHERE BRANCH.SYSTEMBRANCH=True;"
            PRNmc.ParamValue2 = GetFieldValue(Qry)
            PRNmc.ParamValue3 = cboBranch.Text
            PRNmc.ParamValue4 = txtAddress.Text
            'PRNmc.ParamValue5 = txtPO.Text
            PRNmc.ParamValue6 = txtRefNo.Text
            PRNmc.ParamValue7 = txtDate.Text
            PRNmc.ParamValue8 = IIf(chkFinal.Value <> 0, "FINAL", "PENDING")
            PRNmc.ParamValue9 = txtRemarks.Text
            PRNmc.ParamValue10 = txtPrepBy.Text
            PRNmc.ParamValue11 = txtAprvBy.Text
            PRNmc.ParamValue12 = "PACKING NO."
            
'            If chkFinal.Value = 1 Then
'                PRNmc.ParamValue12 = "PACKING NO."
'            Else
'                PRNmc.ParamValue12 = "LOCATION"
'            End If
            PRNmc.Show
            'Me.Visible = False
            
        Case 4  'Exit
            'ProperExit = True
            
            Unload Me
            MotoRep.imgBG.Visible = True
            MotoRep.ToolbarButton.Visible = True
            'LogIn.Show
        Case 5  'Save
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            myMsg = ValidateTransfer
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save Parts Transfer?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Save Spare Parts Transfer
            
            'HIDE THIS IF 3S ANG PROGRAM
            '---
            'If IsAvail = False Then
            '    MsgBox (OrdRows & " Rows No Stock Availability ... in Check mark Rows."), vbCritical, "Available Stock"
            '    Exit Sub
            'End If
            '---
            
            Call SaveTransfer
            'If chkFinal.Value = 1 Then
            If chkFinal.Value = 1 Then Call FinalSPTransfer
            MsgBox "Spareparts Transfer Successfully Saved!", vbInformation + vbOKOnly, "SAVE"
            Call LockData(True)
            Call ButtonSaveCancel(False)
            chkFinal.Enabled = False
            
            cmdCheck.Enabled = True
            cmdCheckUp.Enabled = True
        Case 6  'Cancel
            txtFlx.Visible = False
            txtPartNo.Visible = False
            flxParts.Visible = False
            Ans = MsgBox("Do you want to cancel Transfer?", vbQuestion + vbYesNo)
            If Ans = vbNo Then Exit Sub
            Call ButtonSaveCancel(False)
            Call ClearEntries
            Call LockData(False)
            chkFinal.Enabled = False
    End Select
    
End Sub

Private Sub cmdTrans_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPTRANSFER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
End Sub


Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xRow As Integer, xBal As Integer, i As Integer
    
    'Clear FlxStock
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyID & ")) ORDER BY SPITR.TRANSDATE;"
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
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
            'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
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
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
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
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(StartDate, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(EndDate, "yyyy-mm-dd") & "';"

    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

Private Sub flxOrd_Click()
    Dim MyDesc As String, PrvW As Integer, CrnW As Integer, MaxW As Integer, MyBO As Integer, MySQ As Integer
    
    ActRow = flxOrd.Row
    ActCol = flxOrd.Col
    
    If ActCol = 8 Then
        With txtcorr
            If flxOrd.TextMatrix(ActRow, 4) <> "" Then
                If flxOrd.TextMatrix(ActRow, 8) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 8)
                .Top = flxOrd.CellTop + flxOrd.Top
                .Left = flxOrd.CellLeft + flxOrd.Left
                .Width = flxOrd.CellWidth
                .Text = flxOrd.Text
                .Visible = True
                .SelStart = 0
                .ZOrder
                .SetFocus
            End If
        End With
    End If
    'Call cost History
    Call GetPriceHist
    
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
        Case 6  'Packing
            With txtPacking
                If flxOrd.TextMatrix(ActRow, 4) <> "" Then
                    If flxOrd.TextMatrix(ActRow, 6) <> "" Then .Tag = flxOrd.TextMatrix(ActRow, 6)
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
Private Sub flxOrd_EnterCell()
    Call GetPriceHist
    lblSupplier.Caption = SupplierNm(flxOrd.TextMatrix(flxOrd.Row, 4))
End Sub
Private Sub flxOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxOrd_Click
End Sub
Private Sub flxOrd_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPTRANSFER"
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
    Dim PctDisc As Single, myVat As Single, i As Integer
    Dim DiscPrice As Currency, TtlPrice As Currency, MyCost As Currency, CostVat As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        'flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
        'flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
        'flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
        'flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
        'flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
        'flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
        
        flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
        flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 6) 'Location
        flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 7) 'Applic Model
        flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 1) 'SPid
        flxOrd.TextMatrix(ActRow, 13) = .TextMatrix(.Row, 8) 'Max Discount
        
        If Val(txtVat.Text) > 0 Then
            MyCost = Val(.TextMatrix(.Row, 9)) 'COST
            CostVat = MyCost * (1 + (Val(txtVat.Text) / 100))
            flxOrd.TextMatrix(ActRow, 7) = CostVat 'COST
        Else
            flxOrd.TextMatrix(ActRow, 7) = .TextMatrix(.Row, 9)
        End If
        'Qty and SRP >0
        If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
            TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(.TextMatrix(.Row, 4))
            'flxOrd.TextMatrix(ActRow, 8) = Format(TtlPrice, "#,###,##0.00")
            flxOrd.TextMatrix(ActRow, 9) = Format(TtlPrice, "#,###,##0.00")
        Else
            'flxOrd.TextMatrix(ActRow, 8) = "0.00"
            flxOrd.TextMatrix(ActRow, 9) = "0.00"
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
    Dim DiscPrice As Currency, TtlPrice As Currency, MyCost As Currency, CostVat As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        'flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
        'flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
        'flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
        'flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
        'flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
        'flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
        
        flxOrd.TextMatrix(ActRow, 1) = .TextMatrix(.Row, 5) 'QtyOnHand
        flxOrd.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 3) 'Description
        flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 6) 'Location
        flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 7) 'Applic Model
        flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 1) 'SPid
        flxOrd.TextMatrix(ActRow, 13) = .TextMatrix(.Row, 8) 'Max Discount
        
        If Val(txtVat.Text) > 0 Then
            MyCost = Val(.TextMatrix(.Row, 9)) 'COST
            CostVat = MyCost * (1 + (Val(txtVat.Text) / 100))
            flxOrd.TextMatrix(ActRow, 7) = CostVat 'COST
        Else
            flxOrd.TextMatrix(ActRow, 7) = .TextMatrix(.Row, 9)
        End If
        'Qty and SRP >0
        If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
            TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(.TextMatrix(.Row, 4))
            'flxOrd.TextMatrix(ActRow, 8) = Format(TtlPrice, "#,###,##0.00")
            flxOrd.TextMatrix(ActRow, 9) = Format(TtlPrice, "#,###,##0.00")
            
        Else
            'flxOrd.TextMatrix(ActRow, 8) = "0.00"
            flxOrd.TextMatrix(ActRow, 9) = "0.00"
            
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
            MsgBox "Qty Transferred must not be greater than Qty Available!", vbExclamation + vbOKOnly, "Validation"
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
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPTRANSFER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
    chkFinal.Enabled = False
End Sub
Private Sub Form_Load()
    Dim rsBranch As New ADODB.Recordset
    Dim Qry As String
    
    
    ActRow = 0
    ActCol = 0
    JustScroll = True
    MotoRep.sBar.Panels(6).Text = Me.Caption
    Call ButtonSaveCancel(False)
    txtDate.Text = Format(Date, "mm/dd/yyyy")
    'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='VAT'));"
    'mySQL
    Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='VAT';"
    txtVat.Text = GetFieldValue(Qry)

    With rsBranch
        Qry = "SELECT BRANCH.BRANCHNAME, BRANCH.BRANCHID From BRANCH ORDER BY BRANCH.BRANCHNAME;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboBranch.AddItem !BRANCHNAME
                cboBranch.ItemData(cboBranch.ListCount - 1) = !BRANCHID
                .MoveNext
            Loop
        End If
        .Close
    End With
    chkFinal.Enabled = False

End Sub

Private Sub ButtonSaveCancel(bln As Boolean)
    Dim i As Integer
    
    If bln Then
        For i = 0 To 4
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
       
        
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
        
        
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If ProperExit Then
'        If rsParts.State = adStateOpen Then rsParts.Close
'        ProperExit = False
'    ElseIf JustScroll Then
'        Call cmdTrans_Click(2)
'        ProperExit = False
'    Else
'        Cancel = True
'    End If
'    MotoRep.toolBarButton.Visible = True
'    MotoRep.imgBG.Visible = True
    MotoRep.sBar.Panels(6).Text = ""
End Sub
Private Sub mnuAdd_Click()
    'If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 11) = "" Then
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
        If flxOrd.Rows = 2 Then
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        Else
            flxOrd.RemoveItem ActRow
        End If
        Call ComputeTotal
    End If
End Sub
Public Sub ShowRecord(MyRef As Long)
    
    Dim rsTransfer As New ADODB.Recordset
    Dim rsTransDtl As New ADODB.Recordset
    Dim Qry As String, MyITRef As String
    Dim TtlCost As Currency, i As Integer, MyCost As Currency
    
    
    cmdCheck.Enabled = True
    Call ClearEntries
    'Qry = "SELECT STOCKTRANSFER.* From STOCKTRANSFER WHERE (((STOCKTRANSFER.TRANSFERID)=" & MyRef & "));"
    'mySQL
    
    Qry = "SELECT STOCKTRANSFER.* From STOCKTRANSFER WHERE STOCKTRANSFER.TRANSFERID=" & MyRef & ";"
    
    With rsTransfer
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboBranch.Text = !RECBRANCH
            txtAddress.Text = !Address
            txtSQSERVE.Text = IIf(IsNull(!SQSERVE), "", !SQSERVE)
            txtPO.Text = IIf(IsNull(!PONO), "", !PONO)
            txtRefNo.Text = IIf(Len(str(!TRANSFERID)) >= 4, !TRANSFERID, Format(!TRANSFERID, "0000"))
            txtDate.Text = Format(!TRANSDATE, "MM/DD/YYYY")
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            If !STATUS = "PENDING" Then
                chkFinal.Value = 0
            ElseIf !STATUS = "FINAL" Or !STATUS = "INTRANSIT" Then
                chkFinal.Value = 1
                txtIntransit.Text = ""
                txtIntransit.Text = !STATUS
            ElseIf !STATUS = "CANCELLED" Then
                txtRemarks.Text = "CANCELLED"
            End If
            'txtRemarks.Text = IIf(!Status <> "CANCELLED", !Remarks, "CANCELLED")
            txtPrepBy.Text = !PREPBY
            txtAprvBy.Text = !APRVBY
            txtTotal.Text = Format$(!TotalCost, "#,##0.00")
        End If
        .Close
    End With
    'txtSQSERVE.Text = ""
    
    
    'DISPLAY DETAILS
    If chkFinal.Value = 1 Then
        MyITRef = "TR#" & Val(txtRefNo.Text)
        'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.PRICE, SPAREPARTS.LOCATION, SPITR.SPAREPARTID, SPAREPARTS.APPLICMODEL " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & MyITRef & "') AND ((SPITR.REMARKS)='TRANSFER')) ORDER BY SPITR.ORDER;"
        'mySQL
        'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.PRICE, SPAREPARTS.LOCATION, SPITR.SPAREPARTID, SPAREPARTS.APPLICMODEL " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyITRef & "' AND SPITR.REMARKS='TRANSFER' ORDER BY SPITR.ORDER;"
        If txtIntransit.Text = "INTRANSIT" Then
            
            Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.COST, SPAREPARTS.LOCATION, SPITR.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPITR.POREF " & _
                    "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyITRef & "'" ' ORDER BY SPITR.ORDER;"
        
        Else
            Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.COST, SPAREPARTS.LOCATION, SPITR.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPITR.POREF " & _
                    "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyITRef & "' AND SPITR.REMARKS='TRANSFER'" ' ORDER BY SPITR.ORDER;"
        
        End If
        With rsTransDtl
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxOrd.Rows = 1
                Do While Not .EOF
                    'mySQL
                    'flxOrd.AddItem vbTab & vbTab & !Out & vbTab & IIf(!Out > 1, "Pcs.", "Pc.") & vbTab & !PartNumber & vbTab & !Description & vbTab & !PackingList & vbTab & Format$(!Price, "#,##0.00") & vbTab & Format$(!Price * !Out, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparePartID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel)
                     flxOrd.AddItem vbTab & vbTab & !OUT & vbTab & IIf(!OUT > 1, "Pcs.", "Pc.") & vbTab & !PartNumber & vbTab & !Description & vbTab & !PACKINGLIST & vbTab & Format$(!Cost, "#,##0.00") & vbTab & " " & vbTab & Format$(!Cost * !OUT, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparepartID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & vbTab & vbTab & !POref
                    
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
    
        'Qry = "SELECT SPAREPARTS.QTYONHAND, STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, STOCKTRANSDTL.PACKING, STOCKTRANSDTL.COST, SPAREPARTS.LOCATION, STOCKTRANSDTL.SPAREPARTID, SPAREPARTS.APPLICMODEL " & _
              "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((STOCKTRANSDTL.TRANSFERID) = " & MyRef & ")) ORDER BY STOCKTRANSDTL.SEQNO;"
        
        'mySQL
        Qry = "SELECT SPAREPARTS.QTYONHAND, STOCKTRANSDTL.QTY, STOCKTRANSDTL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, STOCKTRANSDTL.PACKING, STOCKTRANSDTL.COST, SPAREPARTS.LOCATION, STOCKTRANSDTL.SPAREPARTID, SPAREPARTS.APPLICMODEL, STOCKTRANSDTL.POREF " & _
              "FROM STOCKTRANSDTL INNER JOIN SPAREPARTS ON STOCKTRANSDTL.SPAREPARTID = SPAREPARTS.SPAREPARTID Where STOCKTRANSDTL.TRANSFERID = " & MyRef & " ORDER BY STOCKTRANSDTL.SEQNO;"
        
        With rsTransDtl
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxOrd.Rows = 1
                TtlCost = 0
                Do While Not .EOF
                    'flxOrd.AddItem vbTab & !QtyOnHand & vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & !Packing & vbTab & Format$(!Cost, "#,##0.00") & vbTab & Format$(!Cost * !Qty, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparePartID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel)
                    
                    flxOrd.AddItem vbTab & !QtyOnHand & vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & !PACKING & vbTab & Format$(!Cost, "#,##0.00") & vbTab & " " & vbTab & Format$(!Cost * !Qty, "#,##0.00") & vbTab & IIf(IsNull(!Location), "", !Location) & vbTab & !SparepartID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & vbTab & vbTab & !POref
                    TtlCost = TtlCost + (!Qty * !Cost)
                    .MoveNext
                Loop
            End If
            .Close
        End With
        If Val(Replace(txtTotal.Text, ",", "")) <> TtlCost Then
            'LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.TOTALCOST = " & TtlCost & " WHERE (((STOCKTRANSFER.TRANSFERID)=" & Val(txtRefNo.Text) & "));"
            'mySQL
            LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.TOTALCOST = " & TtlCost & " WHERE STOCKTRANSFER.TRANSFERID=" & Val(txtRefNo.Text) & ";"
            
            txtTotal.Text = Format(TtlCost, "#,##0.00")
        End If
    End If
    
    If chkFinal.Value = 1 Then
        MyITRef = "TR#" & Val(txtRefNo.Text)
        TtlCost = 0
        For i = 1 To flxOrd.Rows - 1
            MyCost = Val(Replace(flxOrd.TextMatrix(i, 7), ",", ""))
            If MyCost = 0 Then  'getcost
                'Qry = "SELECT SPAREPARTS.COST From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(flxOrd.TextMatrix(i, 10)) & "));"
                'mySQL
                'Qry = "SELECT SPAREPARTS.COST From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID=" & Val(flxOrd.TextMatrix(i, 10)) & ";"
                Qry = "SELECT SPAREPARTS.COST From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID=" & Val(flxOrd.TextMatrix(i, 11)) & ";"
                
                MyCost = Val(GetFieldValue(Qry))
                If MyCost > 0 Then
                    If Val(txtVat.Text) > 0 Then MyCost = MyCost * (1 + (Val(txtVat.Text) / 100))
                    
                    flxOrd.TextMatrix(i, 7) = Format(MyCost, "#,##0.00") 'COST
                    'flxOrd.TextMatrix(i, 8) = Format(Val(flxOrd.TextMatrix(i, 2)) * MyCost, "#,##0.00")
                    flxOrd.TextMatrix(i, 9) = Format(Val(flxOrd.TextMatrix(i, 2)) * MyCost, "#,##0.00")
                    'LYRcon.Execute "UPDATE SPITR SET SPITR.PRICE = " & MyCost & " WHERE (((SPITR.SPAREPARTID)=" & Val(flxOrd.TextMatrix(i, 10)) & ") AND ((SPITR.REFERENCE)='" & MyITRef & "') AND ((SPITR.REMARKS)='TRANSFER'));"
                    'mySQL
                    LYRcon.Execute "UPDATE SPITR SET SPITR.PRICE = " & MyCost & " WHERE SPITR.SPAREPARTID=" & Val(flxOrd.TextMatrix(i, 11)) & " AND SPITR.REFERENCE='" & MyITRef & "' AND SPITR.REMARKS='TRANSFER';"
                End If
            End If
            TtlCost = TtlCost + (MyCost * Val(flxOrd.TextMatrix(i, 2)))
        Next i
        If Val(Replace(txtTotal.Text, ",", "")) <> TtlCost Then
            'LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.TOTALCOST = " & TtlCost & " WHERE (((STOCKTRANSFER.TRANSFERID)=" & Val(txtRefNo.Text) & "));"
            'mySQL
            LYRcon.Execute "UPDATE STOCKTRANSFER SET STOCKTRANSFER.TOTALCOST = " & TtlCost & " WHERE STOCKTRANSFER.TRANSFERID=" & Val(txtRefNo.Text) & ";"
            
            txtTotal.Text = Format(TtlCost, "#,##0.00")
        End If
    End If
    
End Sub
Private Sub FinalSPTransfer()
    Dim i As Integer, MyRef As String
    Dim Qry As String
    
    On Error GoTo LocalErr
        
        Qry = "SELECT SPORDERTAKING.DATECONFIRM From SPORDERTAKING Where SPORDERTAKING.SQSERVE = " & Val(txtSQSERVE.Text) & ";"
        
        If GetFieldValue(Qry) = "" And Val(txtSQSERVE.Text) > 0 Then
            MsgBox ("Order Not yet Confirm"), vbExclamation, "Not Confirm"
            Exit Sub
        End If
            DateNow = Format(Now, "yyyy-mm-dd")
            LYRcon.Execute "UPDATE stocktransfer SET DATEFINAL = '" & DateNow & "' WHERE stocktransfer.TRANSFERID = " & Val(txtRefNo.Text) & ";"

    With flxOrd
        PbaR.Value = 0
        PbaR.Max = .Rows - 1
        PbaR.Visible = True
        For i = 1 To flxOrd.Rows - 1
            eMsg = "[673] Adding Spareparts Transfer Details @ row " & i
            MyRef = "TR#" & Val(txtRefNo.Text)
            MyDate = Format(Now, "mm/dd/yyyy hh:mm:ss")
            wCost = GetFieldValue("SELECT spareparts.cost From spareparts Where spareparts.SPAREPARTID = " & Val(.TextMatrix(i, 11)) & ";")
                                    
            'LYRcon.Execute "INSERT INTO SPITR (REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, OUT, PRICE, REMARKS, PACKINGLIST, SOURCE, UPDATEDBY, LASTUPDATE, [ORDER]) " & _
                           "VALUES ('" & MyRef & "', 'OUT', #" & txtDate.Text & "#, " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", 'TRANSFER', '" & .TextMatrix(i, 6) & "', '" & Replace(cboBranch.Text, "'", "`") & "', " & UserID & ", #" & Now & "#, " & i & ");"
            'mySQL
            
            LYRcon.Execute "INSERT INTO SPITR (REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, `OUT`, PRICE, REMARKS, PACKINGLIST, SOURCE, UPDATEDBY, LASTUPDATE, `ORDER`, POREF, COST) " & _
            "VALUES ('" & MyRef & "', 'OUT', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 11)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", 'TRANSFER', '" & .TextMatrix(i, 6) & "', '" & Replace(cboBranch.Text, "'", "`") & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & i & ", '" & .TextMatrix(i, 15) & "', " & wCost & ");"
            
            eMsg = "[679] Error updating SP Qty on hand @ row " & i
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]-" & Val(.TextMatrix(i, 2)) & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = #" & Now & "# " & _
                           "WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(.TextMatrix(i, 10)) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 2)) & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                           "WHERE SPAREPARTS.SPAREPARTID=" & Val(.TextMatrix(i, 11)) & ";"
            
            PbaR.Value = PbaR.Value + 1
            DoEvents
        Next i
        PbaR.Visible = False
    End With
    
    If Val(txtSQSERVE) > 0 Then Call UpdateBranchPO
    
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"

    
End Sub
Private Sub SaveTransfer()
    Dim i As Integer, myTtl As Single, TransID As Integer
    Dim eMsg As String, Qry As String, MyBranch As String, MyDate As Date, MyStat As String
                                  
    'Saving Table Spare Parts Invoice
    On Error GoTo LocalErr
    myTtl = Val(Replace(txtTotal.Text, ",", ""))
    MyBranch = Replace(cboBranch.Text, "'", "`")
    'MyDate = Format(Now, "mm/dd/yyyy hh:mm:ss")
    'mySQL
    MyDate = Format(Now, "yyyy-mm-dd hh:mm:ss")
    MyStat = IIf(chkFinal.Value <> 0, "FINAL", "PENDING")
    If AddNew Then
        eMsg = "Saving New Record for SP Transfer"
        'LYRcon.Execute "INSERT INTO STOCKTRANSFER (RECBRANCH, TRANSDATE, ADDRESS, PONO, PREPBY, APRVBY, TOTALCOST, STATUS, REMARKS, UPDATEDBY, LASTUPDATE) " & _
        "VALUES ('" & MyBranch & "', #" & txtDate.Text & "#, '" & Replace(txtAddress.Text, "'", "`") & "', '" & txtPO.Text & "', '" & txtPrepBy.Text & "', '" & txtAprvBy.Text & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", '" & MyStat & "', '" & txtRemarks.Text & "', " & UserID & ", #" & MyDate & "#);"
        'mySQL
        LYRcon.Execute "INSERT INTO STOCKTRANSFER (RECBRANCH, TRANSDATE, ADDRESS, PONO, PREPBY, APRVBY, TOTALCOST, STATUS, REMARKS, UPDATEDBY, LASTUPDATE, SQSERVE) " & _
        "VALUES ('" & MyBranch & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & Replace(txtAddress.Text, "'", "`") & "', '" & txtPO.Text & "', '" & txtPrepBy.Text & "', '" & txtAprvBy.Text & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", '" & MyStat & "', '" & txtRemarks.Text & "', " & UserID & ", '" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "', " & Val(txtSQSERVE) & ");"
        
        'Qry = "SELECT STOCKTRANSFER.TRANSFERID From STOCKTRANSFER " & _
              "WHERE (((STOCKTRANSFER.RECBRANCH)='" & MyBranch & "') AND ((STOCKTRANSFER.LASTUPDATE)=#" & MyDate & "#));"
        'mySQL
        Qry = "SELECT STOCKTRANSFER.TRANSFERID From STOCKTRANSFER " & _
              "WHERE STOCKTRANSFER.RECBRANCH='" & MyBranch & "' AND STOCKTRANSFER.LASTUPDATE='" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "';"

        TransID = GetFieldValue(Qry)
        txtRefNo.Text = IIf(Len(str(TransID)) >= 4, TransID, Format(TransID, "0000"))
    Else
        eMsg = "Updating Record for SP Sales Order"
        TransID = Val(txtRefNo.Text)
        'LYRcon.Execute "UPDATE STOCKTRANSFER SET RECBRANCH = '" & MyBranch & "', TRANSDATE = #" & txtDate.Text & "#, ADDRESS = '" & txtAddress.Text & "', PONO = '" & txtPO.Text & "', PREPBY = '" & txtPrepBy.Text & "', APRVBY = '" & txtAprvBy.Text & "', TOTALCOST = " & Val(Replace(txtTotal.Text, ",", "")) & ", STATUS = '" & MyStat & "', REMARKS = '" & txtRemarks.Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = #" & MyDate & "# " & _
        "WHERE (STOCKTRANSFER.TRANSFERID =" & TransID & ");"
        'mySQL
        LYRcon.Execute "UPDATE STOCKTRANSFER SET RECBRANCH = '" & MyBranch & "', TRANSDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "', ADDRESS = '" & txtAddress.Text & "', PONO = '" & txtPO.Text & "', PREPBY = '" & txtPrepBy.Text & "', APRVBY = '" & txtAprvBy.Text & "', TOTALCOST = " & Val(Replace(txtTotal.Text, ",", "")) & ", STATUS = '" & MyStat & "', REMARKS = '" & txtRemarks.Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "' " & _
        "WHERE (STOCKTRANSFER.TRANSFERID =" & TransID & ");"
    End If
    'Saving Table STOCK TRANSFER Details
    If Not AddNew Then
        'LYRcon.Execute "DELETE stocktransdtl.* From stocktransdtl " & _
                       "WHERE (((stocktransdtl.TRANSFERID)=" & TransID & "));"
        'mySQL
        LYRcon.Execute "DELETE From stocktransdtl " & _
                       "WHERE stocktransdtl.TRANSFERID=" & TransID & ";"
        DoEvents
    End If
    With flxOrd
        PbaR.Value = 0
        PbaR.Max = .Rows - 1
        PbaR.Visible = True
        For i = 1 To flxOrd.Rows - 1
            eMsg = "Line 612: Adding Spareparts Transfer Details @ row " & i
            
            LYRcon.Execute "INSERT INTO stocktransdtl (TRANSFERID, SPAREPARTID, QTY, UNIT, PACKING, COST, SEQNO, POREF ) " & _
                               "VALUES (" & TransID & ", " & Val(.TextMatrix(i, 11)) & ", " & Val(.TextMatrix(i, 2)) & ", '" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 6) & "', " & Val(Replace(.TextMatrix(i, 7), ",", "")) & ", " & i & ", '" & .TextMatrix(i, 15) & "');"
            PbaR.Value = PbaR.Value + 1
            DoEvents
        Next i
        PbaR.Visible = False
    End With
    
    Exit Sub
LocalErr:
    MsgBox "Error " & eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Function ValidateTransfer() As String
    Dim MyErr As String, Qry As String
    Dim i As Integer
    
    MyErr = ""
    If cboBranch.Text = "" Then MyErr = "Please specify the Branch were to transfer the stocks."
    If txtAddress.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Address"
    If Not IsDate(txtDate.Text) Then MyErr = MyErr & vbCrLf & "Invalid entry on field Date"
    If txtPrepBy.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Prepared By"
    If txtAprvBy.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Approved By"
    'If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" And flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" And flxOrd.TextMatrix(flxOrd.Rows - 1, 11) = "" Then
        
        If flxOrd.Rows = 2 Then
            flxOrd.Rows = 1
            flxOrd.Rows = 2
        Else
            flxOrd.RemoveItem flxOrd.Rows - 1
        End If
    End If
    
    'If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 10) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row of Ordered Spare Parts"
    If flxOrd.TextMatrix(flxOrd.Rows - 1, 2) = "" Or flxOrd.TextMatrix(flxOrd.Rows - 1, 11) = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on Last Row of Ordered Spare Parts"
    For i = 1 To flxOrd.Rows - 1
        If Val(flxOrd.TextMatrix(i, 2)) = 0 Then MyErr = MyErr & vbCrLf & "Ordered Spare Parts Sales @ Row " & i & " no Qty.!"
        'If Val(flxOrd.TextMatrix(i, 7)) = 0 Then MyErr = MyErr & vbCrLf & "Ordered Spare Parts Sales @ Row " & i & " no Price!"
    Next i
    ValidateTransfer = MyErr
End Function

Private Sub RXResizer1_ResizeComplete()
    
    With flxOrd
        If showcost Then
            'old
            '.ColWidth(0) = 300
            '.ColWidth(1) = (.Width - 300) * 0.05
            '.ColWidth(2) = (.Width - 300) * 0.07
            '.ColWidth(3) = (.Width - 300) * 0.07
            '.ColWidth(4) = (.Width - 300) * 0.2
            '.ColWidth(5) = (.Width - 300) * 0.3 + (.Width - 300) * 0.1
            '.ColWidth(6) = 0
            '.ColWidth(7) = (.Width - 300) * 0.1
            
            '.TextMatrix(0, 8) = ""
            '.ColWidth(8) = 0
            '.ColWidth(9) = (.Width - 300) * 0.1
            '.ColWidth(10) = 0
            '.ColWidth(11) = 0
            '.ColWidth(12) = 0
            '.ColWidth(13) = 0
            '.ColWidth(14) = 0
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.05
            .ColWidth(2) = (.Width - 300) * 0.06
            .ColWidth(3) = (.Width - 300) * 0.06
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.25 + (.Width - 300) * 0.1
            .ColWidth(6) = 0
            .ColWidth(7) = (.Width - 300) * 0.13
            
            .TextMatrix(0, 8) = ""
            .ColWidth(8) = 0
            .ColWidth(9) = (.Width - 300) * 0.1
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = (.Width - 300) * 0.1
            
            txtTotal.Left = flxOrd.Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7)
            txtTotal.Width = .ColWidth(9)
            txtTotal.Top = flxOrd.Top + flxOrd.Height + 50
            Frame2.Top = txtTotal.Top + txtTotal.Height + 50
            lblTotal.Left = txtTotal.Left - (lblTotal.Width + 50)
            lblTotal.Top = txtTotal.Top
            lblTotal.FontSize = txtTotal.FontSize
            lblTotal.Visible = True
            txtTotal.Visible = True
            flxOld.Visible = True
            cmdCheck.Visible = True
            cmdCheckUp.Visible = True
            frmSup.Visible = True
            
         Else
            'Old
            '.ColWidth(0) = 300
            '.ColWidth(1) = (.Width - 300) * 0.05
            '.ColWidth(2) = (.Width - 300) * 0.08
            '.ColWidth(3) = (.Width - 300) * 0.08
            '.ColWidth(4) = (.Width - 300) * 0.2
            '.ColWidth(5) = ((.Width - 300) * 0.6)
            '.ColWidth(6) = (.Width - 300) * 0.2
            '.ColWidth(7) = 0
            '.TextMatrix(0, 8) = ""
            '.ColWidth(8) = 0
            '.ColWidth(9) = (.Width - 300) * 0.12
            '.ColWidth(10) = 0
            '.ColWidth(11) = 0
            '.ColWidth(12) = 0
            '.ColWidth(13) = 0
            '.ColWidth(14) = 0
            
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.05
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.08
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = ((.Width - 300) * 0.4)
            .ColWidth(6) = 0
            .ColWidth(7) = (.Width - 300) * 0.2
            .TextMatrix(0, 8) = ""
            .ColWidth(8) = 0
            .ColWidth(9) = (.Width - 300) * 0.12
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = (.Width - 300) * 0.2
            flxOld.Visible = False
            cmdCardex.Left = (cmdAdvFind.Left) - cmdCardex.Width
            'flxOrd.Width = flxOrd.Width + flxOld.Width
            flxOrd.Width = Frame1.Width
         End If
        .Col = 6: .Row = 1
        flxParts.ColWidth(2) = .ColWidth(4)
    End With
    
    If ActRow > 0 And ActCol > 0 And Not AddNew Then
        flxOrd.Row = ActRow
        flxOrd.Col = ActCol
        txtPartNo.Top = flxOrd.CellTop + flxOrd.Top
        txtPartNo.Left = flxOrd.CellLeft + flxOrd.Left
        If txtPartNo.Visible = True Then txtPartNo.Width = flxOrd.CellWidth
    End If
    With flxParts
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
    End With
   ' lblHeading(1).Left = Frame1.Left + Frame1.Width - lblHeading(1).Width
'    lblHeading(0).Left = Frame1.Left + Frame1.Width - lblHeading(0).Width - 30
    lblHeading(4).Top = flxOrd.Top - lblHeading(4).Height + 30
    lblHeading(5).Top = flxOrd.Top - lblHeading(5).Height
    lblHeading(4).Left = flxOrd.Left
    lblHeading(5).Left = flxOrd.Left + 30
End Sub

Private Sub ReloadSpareParts()
    Dim Qry As String
    
    'Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.'' ,SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST, SPAREPARTS.REMARKS " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION,SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.COST, SPAREPARTS.REMARKS " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    
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
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then flxOrd.SetFocus
End Sub
Private Sub txtcorr_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789."
    If ActCol = 8 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
        If KeyAscii = 13 Then
            flxOrd.TextMatrix(ActRow, 8) = txtcorr.Text
            txtcorr.Visible = False
            flxOrd.SetFocus
            flxOrd.ColSel = 8
            flxOrd.RowSel = ActRow
            Call ComputeTotal
        End If
    End If
End Sub
Private Sub txtcorr_LostFocus()
    txtcorr.Visible = False
End Sub
Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPTRANSFER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
    
End Sub

Private Sub txtPacking_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtPacking.Text = Empty
        flxOrd.TextMatrix(ActRow, 6) = IIf(txtPacking.Tag = "", "", txtPacking.Tag)
        txtPacking.Tag = "skiplostfocus"
        flxOrd.SetFocus
        flxOrd.Col = 2
        txtPacking.Visible = False
    End If
End Sub

Private Sub txtPacking_KeyPress(KeyAscii As Integer)
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call txtPacking_LostFocus
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
Private Sub txtPacking_LostFocus()
    
    If txtPacking.Tag = "skiplostfocus" Then
        txtPacking.Tag = ""
        Exit Sub
    End If
    flxOrd.TextMatrix(ActRow, 6) = txtPacking.Text
    txtPacking.Visible = False
End Sub
Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub
Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPTRANSFER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
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
        txtRefNo.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        flxOrd.SetFocus
        flxOrd.Col = 4  'partnumber
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789/-"
    If ActCol = 1 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If

    If KeyAscii = 13 Then
        flxOrd.SetFocus
        flxOrd.Col = 4  'partnumber
    End If
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If ActCol = 2 Then
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
        'If Val(txtFlx.Text) > Val(.TextMatrix(ActRow, 1)) And .TextMatrix(ActRow, 10) <> "" Then
        If Val(txtFlx.Text) > Val(.TextMatrix(ActRow, 1)) And .TextMatrix(ActRow, 11) <> "" Then
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

Private Sub TOT_PENDINGORDER(SP_ORDERDATE As String, SP_ID As Long)
    Dim rsPOrder As New ADODB.Recordset
    Dim QRY_ORDER As String, PENDINGORDER As Long, S_Order As Long, s_Col As Long

    QRY_ORDER = "SELECT SUM(spsorderdetail.QTY) AS TOTALQTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO " & _
        " From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
        " Where spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.INVOICENO IS NULL Group By " & _
        " spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO IS NULL;"
                
    PENDINGORDER = Val(GetFieldValue(QRY_ORDER))
    txtTotPen.Text = PENDINGORDER
End Sub

Public Sub ComputeTotal()
    Dim i As Integer
    Dim TtlPrice As Currency, GrndTtl As Currency
     
    
    With flxOrd
        GrndTtl = 0
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 2)) > 0 And .TextMatrix(i, 7) <> "" Then
                If Val(.TextMatrix(i, 8)) > 0 Then
                    TtlPrice = Val(.TextMatrix(i, 2)) * Val(Replace(.TextMatrix(i, 8), ",", ""))
                Else
                    TtlPrice = Val(.TextMatrix(i, 2)) * Val(Replace(.TextMatrix(i, 7), ",", ""))
                End If
                GrndTtl = GrndTtl + TtlPrice
                
                '.TextMatrix(i, 8) = Format$(TtlPrice, "#,##0.00")
                .TextMatrix(i, 9) = Format$(TtlPrice, "#,##0.00")
                
            End If
            
            If InStr(.TextMatrix(i, 5), "ENGINE OIL") > 0 Or InStr(.TextMatrix(i, 5), "OIL 4T") > 0 Then
                .TextMatrix(i, 3) = IIf(Val(.TextMatrix(i, 2)) > 1, "Qrts.", "Qrt.") & IIf(InStr(.TextMatrix(i, 3), "*") > 0, "*", "")
            Else
                flxOrd.TextMatrix(i, 3) = IIf(Val(flxOrd.TextMatrix(i, 2)) > 1, "Pcs.", "Pc.") & IIf(InStr(.TextMatrix(i, 3), "*") > 0, "*", "")
            End If

        Next i
        txtTotal.Text = Format$(GrndTtl, "#,##0.00")
    End With
            
End Sub

Private Sub txtPO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxOrd.SetFocus
        flxOrd.Col = 4
        flxOrd.Row = 1
    End If
End Sub

Private Sub txtRefNo_KeyUp(KeyCode As Integer, Shift As Integer)
    SPfind.MySource = "SPTRANSFER"
    If KeyCode = vbKeyF12 Then SPfind.Show vbModal
End Sub

Private Sub txtPartNo_Change()
    Dim TtlPrice As Currency, DiscPrice As Currency
    Dim PctDisc As Single, myVat As Single, MyCost As Single
    
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
            'Change Code
            'flxOrd.TextMatrix(ActRow, 9) = .TextMatrix(.Row, 6) 'Location
            'flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 7) 'Applic Model
            'flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 1) 'SPid
            'flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 8) 'Max Discount
            
            flxOrd.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 6) 'Location
            flxOrd.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 7) 'Applic Model
            flxOrd.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 1) 'SPid
            flxOrd.TextMatrix(ActRow, 13) = .TextMatrix(.Row, 8) 'Max Discount
            
            'If Val(txtVat.Text) > 0 Then
            '    MyCost = Val(.TextMatrix(.Row, 9)) 'COST
            '    CostVat = MyCost * (1 + (Val(txtVat.Text) / 100))
            '    flxOrd.TextMatrix(ActRow, 7) = CostVat 'COST
            'Else
            '    flxOrd.TextMatrix(ActRow, 7) = .TextMatrix(.Row, 9)
            'End If
            
            'Qty and SRP >0
            'If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
            '    TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(.TextMatrix(.Row, 4))
            '    flxOrd.TextMatrix(ActRow, 8) = Format(TtlPrice, "#,###,##0.00")
            'Else
            '    flxOrd.TextMatrix(ActRow, 8) = "0.00"
            'End If
            
            If Val(txtVat.Text) > 0 Then
                MyCost = Val(.TextMatrix(.Row, 9)) 'COST
                CostVat = MyCost * (1 + (Val(txtVat.Text) / 100))
                flxOrd.TextMatrix(ActRow, 7) = CostVat 'COST
            Else
                flxOrd.TextMatrix(ActRow, 7) = .TextMatrix(.Row, 9)
            End If
            'Qty and SRP >0
            If Val(flxOrd.TextMatrix(ActRow, 2)) > 0 Then
                TtlPrice = Val(flxOrd.TextMatrix(ActRow, 2)) * Val(.TextMatrix(.Row, 4))
                flxOrd.TextMatrix(ActRow, 9) = Format(TtlPrice, "#,###,##0.00")
            Else
                flxOrd.TextMatrix(ActRow, 9) = "0.00"
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
            'rsParts.Requery
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
        Call ReloadSpareParts
    Else
        flxParts.Visible = False
        txtPartNo.Visible = False
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Private Sub txtPartNo_GotFocus()
    If rsParts.State = adStateClosed Then Call ReloadSpareParts
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
    Dim i As Integer
    
    If txtPartNo.Tag = "skiplostfocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    'If flxOrd.TextMatrix(ActRow, 10) = "" Or txtPartNo.Text = "" Then
    If flxOrd.TextMatrix(ActRow, 11) = "" Or txtPartNo.Text = "" Then
        
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

Private Sub txtrefno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtAddress.SetFocus
    ElseIf KeyCode = vbKeyDown Then
        txtDate.SetFocus
    End If
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDate.SetFocus
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
        'If Val(.TextMatrix(ActRow, 14)) = 0 Then
        If Val(.TextMatrix(ActRow, 15)) = 0 Then
        
            MsgBox "No available COST for this item." & vbCrLf & "Please inform your Manager.", vbOKOnly + vbExclamation, "Zero Cost"
            flxOrd.TextMatrix(ActRow, 7) = ""
            txtPrice.Visible = False
            Call ComputeTotal
            Exit Sub
        Else
            myVat = IIf(Val(txtVat.Text) = 0, 1.12, 1 + (Val(txtVat.Text) / 100))
            'AmtDsc = Val(.TextMatrix(ActRow, 14)) * myVat
            AmtDsc = Val(.TextMatrix(ActRow, 15)) * myVat
            If Val(Replace(txtPrice.Text, ",", "")) < Format(AmtDsc, "###0.00") Then
                MsgBox "Invalid Price entered!" & vbCrLf & "Please follow the correct Price", vbOKOnly + vbExclamation, "Invalid Price"
                'flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 13), "#,##0.00")
                flxOrd.TextMatrix(ActRow, 7) = Format(flxOrd.TextMatrix(ActRow, 14), "#,##0.00")
                flxOrd.TextMatrix(ActRow, 6) = ""
                txtPrice.Visible = False
                Call ComputeTotal
                Exit Sub
            End If
        End If
        'LessAmt = Val(.TextMatrix(ActRow, 13)) * (Val(.TextMatrix(ActRow, 12)) / 100)
        'AmtDsc = Val(.TextMatrix(ActRow, 13)) - LessAmt
        
        LessAmt = Val(.TextMatrix(ActRow, 14)) * (Val(.TextMatrix(ActRow, 13)) / 100)
        AmtDsc = Val(.TextMatrix(ActRow, 14)) - LessAmt
        
        If UserType <> "Administrator" Then
            If Val(Replace(txtPrice.Text, ",", "")) < AmtDsc Then
                MsgBox "Price must not be less than maximum discounted Price: " & Format(AmtDsc, "#,##0.00") & "!", vbCritical + vbOKOnly, "Validation"
                txtPrice.Tag = "XcdQty"
                txtPrice.SelStart = 0
                txtPrice.SelLength = Len(txtPrice.Text)
                txtPrice.SetFocus
                Exit Sub
            End If
        End If
        'If Val(Replace(txtPrice.Text, ",", "")) < Val(.TextMatrix(ActRow, 13)) Then
        'mySQL
        If Val(Replace(txtPrice.Text, ",", "")) < Val(.TextMatrix(ActRow, 14)) Then
            'Pctd = (Val(.TextMatrix(ActRow, 13)) - Val(Replace(txtPrice.Text, ",", ""))) / Val(.TextMatrix(ActRow, 13))
            Pctd = (Val(.TextMatrix(ActRow, 14)) - Val(Replace(txtPrice.Text, ",", ""))) / Val(.TextMatrix(ActRow, 14))
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
Private Sub DisplayCostHistory(MyHis As String)
    'Dim pHis As Variant
    'Dim pRow As Variant
    'Dim i As Integer
    
    'pHis = Split(MyHis, "*")
    'With flxOld
    '    .Rows = 1
    '    For i = LBound(pHis) To UBound(pHis)
    '        If pHis(i) <> "" Then
    '            If InStr(pHis(i), "^") > 0 Then
    '                pRow = Split(pHis(i), "^")
    '                .AddItem pRow(0) & vbTab & pRow(1) & vbTab & pRow(2) & vbTab & pRow(3)
    '            Else
    '                .AddItem pHis(i)
    '            End If
    '        End If
    '    Next i
    'End With
    
    'New Code
    Dim pHis As Variant
    Dim pRow As Variant
    Dim i As Integer
    
    pHis = Split(MyHis, "*")
    With flxOld
        .Rows = 1
        For i = LBound(pHis) To UBound(pHis)
            If pHis(i) <> "" Then
                If InStr(pHis(i), "^") > 0 Then
                    pRow = Split(pHis(i), "^")
                    If UBound(pRow) = 2 Then
                        pHis(i) = pHis(i) & "^0"
                        pRow = Split(pHis(i), "^")
                    End If
                    .AddItem pRow(0) & vbTab & pRow(1) & vbTab & pRow(2) & vbTab & pRow(3)
                Else
                    .AddItem pHis(i)
                End If
            End If
        Next i
    End With
End Sub
Private Sub GetPriceHist()
Dim SPCostHist As String
Dim sql As String
    
    sql = "SELECT spareparts.PRICEHIST,spareparts.PARTNUMBER From spareparts Where spareparts.PARTNUMBER = '" & flxOrd.TextMatrix(flxOrd.Row, 4) & "';"
    SPCostHist = GetFieldValue(sql)
    Call DisplayCostHistory(SPCostHist)
End Sub
'Load this code after loading sa Checking
Private Sub DisplayCostOfSales()

Dim CostHis As Variant, SplRec As Variant, Cst As Integer, InvDate As String
Dim rw As Long, InvDateCol As String, Ldate As Boolean
    
    With flxOld
        For rw = 1 To .Rows - 1
            'If empty SP query
            If .TextMatrix(rw, 1) = "" Then Exit Sub
            'If empty History of spareparts get get actual cost
            If .TextMatrix(rw, 7) = "" Then
               .TextMatrix(rw, 7) = Format$(.TextMatrix(rw, 11), "#,##0.00")
               .TextMatrix(rw, 8) = Val(.TextMatrix(rw, 6)) * Replace(.TextMatrix(rw, 7), ",", "")
                
            End If
            'Split history by Date changed,(*)
            CostHis = Split(.TextMatrix(rw, 7), "*")
            'Getting S.I date of sales
            InvDateCol = Format(.TextMatrix(rw, 4), "yyyy-mm-dd")
            'Boolean flag
            Ldate = False
            For Cst = UBound(CostHis) To LBound(CostHis) Step -1
                If CostHis(Cst) <> "" Then
                    If InStr(CostHis(Cst), "^") > 0 Then
                        'Split history
                        SplRec = Split(CostHis(Cst), "^")
                        'If invoice date greater than date price changed
                        If InvDateCol >= Format(SplRec(0), "yyyy-mm-dd") And SplRec(2) <> "0" And Ldate = False Then
                            'Get actual cost of invoice
                            .TextMatrix(rw, 7) = Format$(SplRec(2), "#,##0.00")
                            .TextMatrix(rw, 8) = Format$(Val(.TextMatrix(rw, 6)) * Val(Replace(.TextMatrix(rw, 7), ",", "")), "#,##0.00")
                             Ldate = True: Exit For
                        'If invoice date greater than date price changed
                        ElseIf InvDateCol < Format(SplRec(0), "yyyy-mm-dd") And SplRec(2) <> "0" And Ldate = False Then
                            Ldate = True: Exit For
                            .TextMatrix(rw, 7) = Format$(SplRec(2), "#,##0.00")
                            .TextMatrix(rw, 8) = Format$(Val(.TextMatrix(rw, 6)) * Val(Replace(.TextMatrix(rw, 7), ",", "")), "#,##0.00")
                        'Zero cost in history retrieve current cost
                        ElseIf SplRec(2) = "0" Then
                           .TextMatrix(rw, 7) = Format$(.TextMatrix(rw, 11), "#,##0.00")
                           .TextMatrix(rw, 8) = Format$(Val(.TextMatrix(rw, 6)) * Replace(.TextMatrix(rw, 7), ",", ""), "#,##0.00")
                            Ldate = True: Exit For
                        End If
                    End If
                End If
            Next Cst
        Next rw
    End With
End Sub
Function SupplierNm(Sname As String) As String
Dim qryS As String
    qryS = "SELECT supplier.SupplierName From supplier INNER JOIN spareparts ON (supplier.SUPPLIERCODE = spareparts.SUPPLIERCODE) " & _
        "Where spareparts.PARTNUMBER = '" & Sname & "';"
    SupplierNm = CStr(GetFieldValue(qryS))
    
End Function
Private Sub LoadBranchPO()
    Dim rsTransPO As New ADODB.Recordset
    Dim PONumber As String, POSQ As Long
    Dim Qry As String, ref As Integer
        If cboBranch.Text = "" Then
            MsgBox ("Select Branch Name"), vbExclamation, "Branch ID"
            Exit Sub
        End If
        txtSQSERVE.Text = ""
        txtSQSERVE.Text = InputBox("Enter Branch ORDER NUMBER: ", "Order Number")
        Qry = "SELECT spordertaking.SQSERVE From spordertaking Where spordertaking.BRANCHID = " & Val(txtBranchID.Text) & " AND spordertaking.SQSERVE= " & Val(txtSQSERVE) & " AND spordertaking.FINALQT=1;"
        POSQ = Val(GetFieldValue(Qry))
        
        If POSQ = 0 Then
            MsgBox ("No P.O Reference Number found in the Records"), vbExclamation, "No P.O"
            Exit Sub
        Else
            'Qry = "SELECT DISTINCT spordrconfirm.REFERENCE From spordrconfirm where spordrconfirm.SQSERVE = " & Val(txtSQSERVE) & " AND spordrconfirm.FINALQT = 1;"
            Qry = "SELECT DISTINCT spordrconfirm.REFERENCE " & _
                  "From spordertaking RIGHT OUTER JOIN spordrconfirm ON (spordertaking.SQSERVE = spordrconfirm.SQSERVE) " & _
                  "WHERE spordertaking.SQSERVE = " & Val(txtSQSERVE.Text) & " AND spordertaking.FINALQT = 1;"
            rsTransPO.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsTransPO.RecordCount > 0 Then
                rsTransPO.MoveFirst
                For ref = 1 To rsTransPO.RecordCount
                    txtPO.Text = txtPO.Text & IIf((txtPO.Text = ""), "", ",") & rsTransPO!REFERENCE
                    rsTransPO.MoveNext
                Next ref
            End If
            rsTransPO.Close
        End If
        
        Qry = "SELECT spordertaking.DATECONFIRM From spordertaking Where spordertaking.BRANCHID = " & Val(txtBranchID.Text) & " AND spordertaking.ORDERIDSQ=" & Val(txtSQSERVE.Text) & " AND spordertaking.FINALQT=0;"
        If IsNull(GetFieldValue(Qry)) Then
            MsgBox ("Order no yet Confirm.."), vbExclamation, "No P.O"
            Exit Sub
        End If
  
        'If POLoaded Then MsgBox ("P.O Already loaded : " & txtPO.Text), vbExclamation, "P.O": txtPO.Text = "": Exit Sub
        
        Qry = ""
        Qry = "SELECT DISTINCT spbrchconfirmdtl.SPAREPARTID, spareparts.QTYONHAND, spbrchconfirmdtl.ORDERTAKING, spareparts.PARTNUMBER, " & _
              "spareparts.DESCRIPTION, spareparts.COST, spbrchconfirmdtl.REFERENCE " & _
              "From spbrchconfirmdtl INNER JOIN spareparts ON (spbrchconfirmdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
              "LEFT OUTER JOIN spordertaking ON (spbrchconfirmdtl.ORDERID = spordertaking.SQSERVE) " & _
              "Where spbrchconfirmdtl.TRANSTYPE = 'ORDER TEMP' AND spbrchconfirmdtl.BRANCHID = " & Val(txtBranchID.Text) & " AND spordertaking.SQSERVE = " & Val(txtSQSERVE.Text) & " Order By spareparts.PARTNUMBER;"
        
        
             
        With rsTransPO
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxOrd.Rows = 1
                Do While Not .EOF
                       flxOrd.AddItem vbTab & !QtyOnHand & vbTab & !ORDERTAKING & vbTab & vbTab & !PartNumber & vbTab & !Description & vbTab & vbTab & !Cost & vbTab & vbTab & vbTab & vbTab & !SparepartID & vbTab & vbTab & vbTab & vbTab & !REFERENCE
                       Call ComputeTotal
                    .MoveNext
                Loop
            End If
            .Close
        End With
End Sub
Private Function POLoaded() As Boolean
    Dim rsPOChk As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT spordertaking.SQSERVE From spordertaking Where spordertaking.SQSERVE = " & Val(txtSQSERVE.Text) & " AND spordertaking.BRANCHID = " & Val(txtBranchID.Text) & " AND spordertaking.FINALQT = 0; "
 
    With rsPOChk
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            POLoaded = False
        Else
            POLoaded = True
        End If
        .Close
    End With
    
End Function
Private Sub UpdateBranchPO()
    Dim Qry As String, pRW As Long
    Dim myLastUpdt As String, POTranstype As String, TotOrder As Long, ServeOrder As Long, TotBackOrder As Long
    Dim OrderQty As Long, SPBackOrder As Long, TotOTaking As Long, bORDERID As Long
    
    With flxOrd
        myLastUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
        POTranstype = "ORDER"
        Qry = "SELECT spordrconfirm.ORDERID From spordrconfirm  INNER JOIN spordertaking ON (spordrconfirm.SQSERVE = spordertaking.SQSERVE) " & _
              "Where spordertaking.SQSERVE = " & Val(txtSQSERVE.Text) & " AND spordrconfirm.BRANCHNAME = '" & cboBranch.Text & "';"
        bORDERID = Val(GetFieldValue(Qry))
        
        For pRW = 1 To .Rows - 1
            
                LYRcon.Execute "UPDATE spbrchconfirmdtl SET TRANSTYPE =  '" & POTranstype & "', ORDERTAKING = " & Val(.TextMatrix(pRW, 2)) & ", PRICE = " & Val(Replace(.TextMatrix(pRW, 7), ",", "")) & ", UPDATEDBY = " & UserID & " " & _
                        "Where spbrchconfirmdtl.REFERENCE =  '" & .TextMatrix(pRW, 15) & "' AND spbrchconfirmdtl.ORDERID =  " & bORDERID & " AND spbrchconfirmdtl.BRANCHID = " & Val(txtBranchID.Text) & " AND spbrchconfirmdtl.SPAREPARTID= " & Val(.TextMatrix(pRW, 11)) & " AND spbrchconfirmdtl.TRANSTYPE= 'ORDER TEMP';"
                DoEvents
                Qry = "SELECT SUM(spbrchconfirmdtl.ORDERTAKING) As TOTORDERTAKING From spbrchconfirmdtl Where spbrchconfirmdtl.BRANCHID = " & Val(txtBranchID.Text) & " AND spbrchconfirmdtl.REFERENCE = '" & .TextMatrix(pRW, 15) & "' AND spbrchconfirmdtl.TRANSTYPE = 'ORDER' AND spbrchconfirmdtl.SPAREPARTID = " & Val(.TextMatrix(pRW, 11)) & ";"
                ServeOrder = Val(GetFieldValue(Qry))
                
                Qry = "SELECT spbranchdtl.`ORDER` From spbranchdtl Where spbranchdtl.REFERENCE =  '" & .TextMatrix(pRW, 15) & "'  AND spbranchdtl.BRANCHNAME = '" & cboBranch.Text & "' AND spbranchdtl.SPAREPARTID = " & Val(.TextMatrix(pRW, 11)) & ";"
                OrderQty = Val(GetFieldValue(Qry))
                
                SPBackOrder = OrderQty - ServeOrder
                TotBackOrder = TotBackOrder + SPBackOrder
                TotOrder = TotOrder + OrderQty
                TotOTaking = TotOTaking + Val(.TextMatrix(pRW, 2))
                
                LYRcon.Execute "UPDATE spbrchconfirmdtl SET BBACKORDER = " & SPBackOrder & _
                       " Where spbrchconfirmdtl.REFERENCE = '" & .TextMatrix(pRW, 15) & "' AND spbrchconfirmdtl.BRANCHID = " & Val(txtBranchID.Text) & " AND spbrchconfirmdtl.ORDERID=" & bORDERID & " AND spbrchconfirmdtl.SPAREPARTID= " & Val(.TextMatrix(pRW, 11)) & ";"
         Next pRW
            
            LYRcon.Execute "Update spordertaking SET CONFIRMNO = '" & txtRefNo.Text & "', QTYTAKING = " & TotOTaking & ", BRANCHBO = " & TotBackOrder & ", TOTALAMOUNT = " & Val(Replace(txtTotal.Text, ",", "")) & ", FINALQT = 1, BOLASTUPDATE = '" & Format(Now, "yyyy-mm-dd") & "', MYPC = '" & ComputerName & "', LASTUPDATE = '" & Format(Now, "yyyy-mm-dd") & "' " & _
            " Where spordertaking.SQSERVE = " & txtSQSERVE & ";"
            
            Call CheckAllocated
       
    End With
 End Sub
Private Sub CheckAllocated()
    Dim rsWAll As New ADODB.Recordset
    Dim QryAll As String, RefB As Long
    Dim QryTake As String
    Dim RefNo As String, TOTAlldQTY As Long, TOTBQty As Long
    
    QryAll = "SELECT DISTINCT spordrconfirm.REFERENCE, spordrconfirm.BRANCHNAME " & _
        " From spordrconfirm INNER JOIN spordertaking ON (spordrconfirm.SQSERVE = spordertaking.SQSERVE)" & _
        " Where spordertaking.SQSERVE = " & Val(txtSQSERVE.Text) & " AND spordrconfirm.BRANCHNAME = '" & cboBranch.Text & "';"
    
    With rsWAll
        If .State = adStateOpen Then .Close
            .Open QryAll, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                RefNo = rsWAll!REFERENCE
                For RefB = 1 To .RecordCount
                    QryTake = "SELECT DISTINCT SUM(spbrchconfirmdtl.ORDERTAKING) As TOTTAKING " & _
                        "From spordertaking RIGHT OUTER JOIN spbrchconfirmdtl ON (spordertaking.SQSERVE = spbrchconfirmdtl.ORDERID) " & _
                        "Where spbrchconfirmdtl.REFERENCE = '" & RefNo & "' AND spbrchconfirmdtl.BRANCHID = " & Val(txtBranchID.Text) & " AND spordertaking.CONFIRMNO IS NOT NULL;"
                    TOTQTY = Val(GetFieldValue(QryTake))
                    
                    QryTake = "SELECT DISTINCT SUM(spbranchdtl.ORDERBAL) TOTBAL From spbranchdtl " & _
                        "Where spbranchdtl.REFERENCE='" & RefNo & "' AND spbranchdtl.BRANCHNAME='" & cboBranch.Text & "';"
                    TOTBQty = Val(GetFieldValue(QryTake))
                    
                    If TOTQTY < TOTBQty Then
                        LYRcon.Execute "Update sporderbranch SET ALLOCATED = 0 Where sporderbranch.REFERENCE = '" & RefNo & "' AND sporderbranch.BRANCHNAME = '" & cboBranch.Text & "';"
                    Else
                        LYRcon.Execute "Update sporderbranch SET ALLOCATED = 1 Where sporderbranch.REFERENCE = '" & RefNo & "' AND sporderbranch.BRANCHNAME = '" & cboBranch.Text & "';"
                    End If
                    .MoveNext
                Next RefB
            End If
    End With
End Sub
Private Function IsAvail() As Boolean
    Dim Qry1 As String, TotPTrans As Long
    
    Dim TotAvl As Long
    Dim nAvl As Boolean
    
    OrdRows = ""
    pOrder = 0
    nAvl = False
    For pOrder = 1 To flxOrd.Rows - 1
       Qry1 = "SELECT SUM(stocktransdtl.QTY) As TOTTRANS From stocktransfer RIGHT OUTER JOIN stocktransdtl ON (stocktransfer.TRANSFERID = stocktransdtl.TRANSFERID) " & _
         " Where stocktransfer.TRANSFERID <> " & Val(txtRefNo) & " AND stocktransfer.`STATUS` = 'PENDING' AND stocktransdtl.SPAREPARTID = " & Val(flxOrd.TextMatrix(pOrder, 11)) & ";"
        TotPTrans = Val(GetFieldValue(Qry1))
        Call TOT_PENDINGORDER(Format(txtDate.Text, "yyyy-mm-dd"), Val(flxOrd.TextMatrix(pOrder, 11)))
        TotAvl = Val(flxOrd.TextMatrix(pOrder, 1)) - Val(txtTotPen.Text) - TotPTrans
        If TotAvl < Val(flxOrd.TextMatrix(pOrder, 2)) Then
                flxOrd.Row = pOrder
                flxOrd.Col = 0
                flxOrd.CellFontName = "Wingdings"
                flxOrd.CellFontSize = 12
                flxOrd.Text = lblChk.Caption
           nAvl = True
           OrdRows = OrdRows & " " & pOrder & " ,"
         End If
    Next pOrder
    
    If nAvl = True Then
          IsAvail = False
    Else
         IsAvail = True
    End If
    
End Function
