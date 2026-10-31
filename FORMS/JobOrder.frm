VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form JobOrder 
   Caption         =   "Job Order"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   ControlBox      =   0   'False
   Icon            =   "JobOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDiscount 
      Height          =   285
      Left            =   6840
      TabIndex        =   81
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFSC 
      Height          =   285
      Left            =   9960
      TabIndex        =   80
      Top             =   9240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSrvcType 
      Height          =   285
      Left            =   8520
      TabIndex        =   79
      Top             =   9240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPctD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      TabIndex        =   78
      Tag             =   "XYF"
      Top             =   6540
      Width           =   855
   End
   Begin VB.TextBox txtGrandTtl 
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
      Height          =   315
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   75
      Tag             =   "XYF"
      Top             =   6840
      Width           =   1545
   End
   Begin VB.TextBox txtDiscAmt 
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
      Height          =   315
      Left            =   12720
      TabIndex        =   74
      Tag             =   "XYF"
      Top             =   6525
      Width           =   1545
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   360
      TabIndex        =   68
      Tag             =   "XYF"
      Top             =   7680
      Width           =   13935
      Begin VB.ComboBox cboPO 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   82
         Tag             =   "XYF"
         Top             =   520
         Width           =   2535
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
         Height          =   615
         Left            =   8040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   71
         Tag             =   "XYF"
         Top             =   180
         Width           =   5775
      End
      Begin VB.ComboBox cboReq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   70
         Tag             =   "XYF"
         Top             =   180
         Width           =   2535
      End
      Begin lvButton.lvButtons_H cmdReq 
         Height          =   330
         Left            =   240
         TabIndex        =   69
         Tag             =   "XYF"
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Caption         =   "REQ."
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPO 
         Height          =   330
         Left            =   240
         TabIndex        =   83
         Tag             =   "XYF"
         Top             =   520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Caption         =   "P.O."
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         cBack           =   -2147483633
      End
      Begin VB.Label Label11 
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
         Left            =   6840
         TabIndex        =   72
         Tag             =   "XYF"
         Top             =   225
         Width           =   1095
      End
   End
   Begin VB.TextBox txtPosMechanic 
      Height          =   285
      Left            =   6360
      TabIndex        =   67
      Top             =   9360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cboEngine 
      Height          =   315
      Left            =   5040
      TabIndex        =   65
      Top             =   10560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid flxAllMC 
      Height          =   1215
      Left            =   240
      TabIndex        =   64
      Top             =   9240
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
   End
   Begin VB.TextBox txtDesignation 
      Height          =   285
      Left            =   8640
      TabIndex        =   60
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPayStart 
      Height          =   285
      Left            =   7560
      TabIndex        =   59
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   6840
      TabIndex        =   58
      Top             =   9600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkOnAcct 
      Caption         =   "ON ACCOUNT"
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
      Left            =   9960
      TabIndex        =   57
      Tag             =   "XYF"
      Top             =   9480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   4080
      TabIndex        =   53
      Top             =   15000
      Width           =   495
   End
   Begin VB.TextBox txtMC 
      Height          =   285
      Left            =   6600
      TabIndex        =   49
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cboUsage 
      Height          =   315
      Left            =   8280
      TabIndex        =   48
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   360
      TabIndex        =   44
      Tag             =   "XYF"
      Top             =   6300
      Width           =   4815
      Begin VB.TextBox txtMechanicID 
         Height          =   285
         Left            =   4800
         TabIndex        =   55
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cboMechanic 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   54
         Tag             =   "XYF"
         Top             =   885
         Width           =   3015
      End
      Begin VB.TextBox txtOR 
         Height          =   330
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   50
         Tag             =   "XYF"
         Top             =   240
         Width           =   3015
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
         Height          =   330
         Left            =   1440
         TabIndex        =   47
         Tag             =   "XYF"
         Top             =   570
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "O.R. Number:"
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
         TabIndex        =   51
         Tag             =   "XYF"
         Top             =   278
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Mechanic:"
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
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   923
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   608
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxMC 
      Height          =   855
      Left            =   360
      TabIndex        =   43
      Tag             =   "XYF"
      Top             =   2835
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   16777215
      FormatString    =   "^MODEL|^ENGINE NO.|^CHASSIS NO.|^PLATE NO.|^TYPE OF USAGE|^KM. READING"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   6240
      TabIndex        =   42
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TtlOutLabor 
      Height          =   285
      Left            =   6960
      TabIndex        =   41
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   6360
      TabIndex        =   40
      Top             =   9600
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1335
      Left            =   360
      TabIndex        =   39
      Top             =   9600
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2355
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   9480
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   9960
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   9480
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   9960
   End
   Begin VB.ComboBox cboJob 
      Height          =   315
      Left            =   8160
      TabIndex        =   38
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   37
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   36
      Tag             =   "XYF"
      Top             =   6240
      Width           =   1545
   End
   Begin MSFlexGridLib.MSFlexGrid flxJob 
      Height          =   2175
      Left            =   360
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   4080
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "^    |< J O B   R E Q U E S T    |^     FSC Number       |>  LABOR  |> DISCOUNT  |>  T O T A L  "
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
      Height          =   2055
      Left            =   360
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   720
      Width           =   14055
      Begin VB.CheckBox chkAddMC 
         Caption         =   "Addtl MC"
         Height          =   255
         Left            =   6240
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin lvButton.lvButtons_H cmdNewMC 
         Height          =   330
         Left            =   5040
         TabIndex        =   62
         Tag             =   "XYF"
         Top             =   1230
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Caption         =   "NEW MC"
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
      Begin VB.CheckBox chkType 
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   330
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   61
         Tag             =   "XY"
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboBrand 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   1230
         Width           =   3615
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
         Height          =   330
         Left            =   1440
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtAcct 
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
         Height          =   330
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   900
         Width           =   1095
      End
      Begin VB.ComboBox cboType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtDown 
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   1238
         Width           =   1695
      End
      Begin VB.TextBox txtBal 
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   1568
         Width           =   1695
      End
      Begin VB.TextBox txtJobNo 
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
         Left            =   12240
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   248
         Width           =   1695
      End
      Begin VB.TextBox txtDateProm 
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
         Left            =   12240
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   908
         Width           =   1695
      End
      Begin VB.TextBox txtDateRec 
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
         Left            =   12240
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   578
         Width           =   1695
      End
      Begin VB.TextBox txtContact 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   900
         Width           =   2175
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
         Height          =   330
         Left            =   1440
         MaxLength       =   120
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   570
         Width           =   4695
      End
      Begin VB.Label lblAvlMC 
         Alignment       =   2  'Center
         Caption         =   "2 MC"
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
         Left            =   5040
         TabIndex        =   66
         Tag             =   "XYF"
         Top             =   1605
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "SERVICE TYPE"
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
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   1598
         Width           =   1275
      End
      Begin VB.Label Label8 
         Caption         =   "ACCOUNT NO."
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
         Left            =   3720
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   938
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "MC BRAND"
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
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   1268
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "DOWN PAYMENT"
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
         Index           =   8
         Left            =   10680
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   1275
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "BALANCE"
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
         Left            =   10680
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   1605
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "JOB ORDER NO."
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
         Left            =   10680
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   285
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "CONTACT NOS."
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
         Left            =   120
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   938
         Width           =   2000
      End
      Begin VB.Label Label2 
         Caption         =   "DATE PROMISED"
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
         Left            =   10680
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   945
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "DATE DEPOSITED"
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
         Left            =   10680
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   615
         Width           =   1995
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   608
         Width           =   2000
      End
      Begin VB.Label Label1 
         Caption         =   "CUSTOMER"
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
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   278
         Width           =   2000
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&New Job Order"
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
      Image           =   "JobOrder.frx":000C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3356
      TabIndex        =   14
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
      Image           =   "JobOrder.frx":05A6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   15
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Job Order"
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
      Image           =   "JobOrder.frx":0B40
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8385
      TabIndex        =   16
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Job Order"
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
      Image           =   "JobOrder.frx":10DA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10065
      TabIndex        =   17
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
      Image           =   "JobOrder.frx":1674
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   11745
      TabIndex        =   18
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
      Image           =   "JobOrder.frx":1C0E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1678
      TabIndex        =   13
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit Job Order"
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
      Image           =   "JobOrder.frx":21A8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   52
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&On-Acct. list"
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
      Image           =   "JobOrder.frx":2742
      cBack           =   -2147483633
   End
   Begin VB.Label lblGrandTtl 
      Alignment       =   1  'Right Justify
      Caption         =   "GRAND TOTAL:"
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
      Left            =   11280
      TabIndex        =   77
      Tag             =   "XYF"
      Top             =   6870
      Width           =   1335
   End
   Begin VB.Label lblDisc 
      Alignment       =   1  'Right Justify
      Caption         =   "LESS DISCOUNT:"
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
      Left            =   11280
      TabIndex        =   76
      Tag             =   "XYF"
      Top             =   6555
      Width           =   1335
   End
   Begin VB.Label lblTotal 
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
      Left            =   11280
      TabIndex        =   73
      Tag             =   "XYF"
      Top             =   6255
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14925
      m_MDI_Height    =   9420
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "JOB ESTIMATED COST"
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
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   3795
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "JOB ORDER"
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
      Height          =   315
      Index           =   2
      Left            =   11970
      TabIndex        =   31
      Top             =   450
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
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   29
      Top             =   450
      Width           =   3000
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
      Height          =   315
      Index           =   1
      Left            =   390
      TabIndex        =   30
      Top             =   480
      Width           =   3000
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "JOB ORDER"
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
      Height          =   315
      Index           =   3
      Left            =   12000
      TabIndex        =   32
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "JOB ESTIMATED COST"
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
      Height          =   315
      Index           =   5
      Left            =   390
      TabIndex        =   34
      Tag             =   "XYF"
      Top             =   3825
      Width           =   2415
   End
   Begin VB.Menu mnuFlx 
      Caption         =   "flxJob Rows"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   " Add Row         Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   " Delete"
      End
   End
End
Attribute VB_Name = "JobOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCustomer As New ADODB.Recordset
Dim rsJobDesc As New ADODB.Recordset
Dim rsMechanic As New ADODB.Recordset

Dim CustSrch As Boolean, JobSrch As Boolean, CustBln As Boolean, MechSrch As Boolean
Dim ActRow As Integer, ActCol As Integer, MCrow As Integer, MCcol As Integer, cusJONO As String
Dim AddNew As Boolean, BtnIndx As Integer

Private Sub LoadCustomer()
    Dim rsSettings As New ADODB.Recordset

    Dim Qry As String
    
    'Qry = "SELECT SETTINGS.DESCRIPTION, SETTINGS.VALUE, SETTINGS.REMARKS " & _
          "From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='JOB ORDER')) ORDER BY SETTINGS.VALUE;"
    'mySQL
    Qry = "SELECT SETTINGS.DESCRIPTION, SETTINGS.VALUE, SETTINGS.REMARKS " & _
          "From SETTINGS WHERE SETTINGS.DESCRIPTION='JOB ORDER' ORDER BY SETTINGS.VALUE;"
    
    With rsSettings
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If .Fields("Remarks") = "Type of usage" Then
                    cboUsage.AddItem .Fields("Value")
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    '***Access
    'Qry = "SELECT MCSERVICELABOR.JOBREQUEST, MCSERVICELABOR.LABORCOST " & _
          "From MCSERVICELABOR ORDER BY MCSERVICELABOR.JOBREQUEST;"
    'mySQL
    Qry = "SELECT MCSERVICELABOR.JOBREQUEST, MCSERVICELABOR.LABORCOST " & _
          "From MCSERVICELABOR ORDER BY MCSERVICELABOR.JOBREQUEST;"
    
    With rsJobDesc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboJob.AddItem UCase$(.Fields(0))
                .MoveNext
            Loop
        End If
    End With
    
    '***Access
    'Qry = "SELECT Employees.EMPID, [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS MECHANIC, Employees.POSITION From Employees " & _
          "WHERE (((Employees.INACTIVE)=False) AND ((Employees.DEPARTMENT)='SERVICE' Or (Employees.DEPARTMENT)='2H')) ORDER BY [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.';"
    'mySQl
    Qry = "SELECT Employees.EMPID, concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS MECHANIC, Employees.`POSITION` From Employees " & _
        "WHERE Employees.DEPARTMENT='2H' or Employees.INACTIVE=0 AND Employees.DEPARTMENT='SERVICE' order by concat(EMPLOYEES.LASTNAME,', ',left(EMPLOYEES.FIRSTNAME,1),'.');"
    With rsMechanic
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboMechanic.AddItem .Fields(1)
                cboMechanic.ItemData(cboMechanic.ListCount - 1) = .Fields(0)
                .MoveNext
            Loop
        End If
    End With
        
End Sub


Public Sub DisplayJobOrder(MyNo As String)
    Dim rsJob As New ADODB.Recordset
    Dim rsDetails As New ADODB.Recordset
    Dim rsReq As New ADODB.Recordset
    Dim Qry As String, sReq As String
    Dim i As Integer, myTtl As Single, myDisc As Single, MyLabor As Single
    
    If MyNo = "" Then
        MsgBox "Unable to process blank Job Order No.", vbExclamation + vbOKOnly, "Empty Job No."
        Exit Sub
    End If
    '***Access
    'Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CONTACTNO, SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.BRAND, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.DOWNPAYMENT, SRVCJOBORDER.BALANCE, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.TYPEOFUSAGE, SRVCJOBORDER.KMREADING, SRVCJOBORDER.ORNO, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.PREPAREDBY, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.CUSTOMERID, SRVCJOBORDER.PAYTYPE, SRVCUSTOMER.DESIGNATION, SRVCJOBORDER.DISCOUNT " & _
          "FROM SRVCUSTOMER INNER JOIN SRVCJOBORDER ON SRVCUSTOMER.ACCOUNTNO = SRVCJOBORDER.ACCOUNTNO WHERE (((SRVCJOBORDER.JOBORDERID)=" & MyNo & "));"
    '***Access
    'mySQl
    Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CONTACTNO, SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.BRAND, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.DOWNPAYMENT, SRVCJOBORDER.BALANCE, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.TYPEOFUSAGE, SRVCJOBORDER.KMREADING, SRVCJOBORDER.ORNO, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.PREPAREDBY, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.CUSTOMERID, SRVCJOBORDER.PAYTYPE, SRVCUSTOMER.DESIGNATION, SRVCJOBORDER.DISCOUNT " & _
          "FROM SRVCUSTOMER INNER JOIN SRVCJOBORDER ON SRVCUSTOMER.ACCOUNTNO = SRVCJOBORDER.ACCOUNTNO WHERE SRVCJOBORDER.JOBORDERID=" & MyNo & ";"
    
    With rsJob
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtJOid.Text = .Fields("JobOrderID")
            CustSrch = True 'avoid display of FlxCust
            txtCustomer.Text = IIf(IsNull(.Fields("Customer")), "", .Fields("Customer"))
            txtAddress.Text = IIf(IsNull(.Fields("Address")), "", .Fields("Address"))
            txtContact.Text = IIf(IsNull(.Fields("ContactNo")), "", .Fields("ContactNo"))
            txtAcct.Text = IIf(IsNull(.Fields("AccountNo")), "", .Fields("AccountNo"))
            cboType.Text = IIf(IsNull(.Fields("JOType")), "", .Fields("JOType"))    'wla na sa table
            cboBrand.Text = IIf(IsNull(.Fields("Brand")), "", .Fields("Brand"))
            txtJobNo.Text = .Fields("JobOrderNo")
            txtDateRec.Text = .Fields("DateReceived")
            txtDateProm.Text = IIf(IsNull(.Fields("DatePromised")), "", .Fields("DatePromised"))
            txtDown.Text = IIf(IsNull(.Fields("DownPayment")), 0, .Fields("DownPayment"))
            txtBal.Text = IIf(IsNull(.Fields("Balance")), 0, .Fields("Balance"))
            flxMC.Rows = 1
            flxMC.AddItem IIf(IsNull(.Fields("Model")), "", .Fields("Model")) & vbTab & IIf(IsNull(.Fields("Engine")), "", .Fields("Engine")) & vbTab & IIf(IsNull(.Fields("Chassis")), "", .Fields("Chassis")) & vbTab & IIf(IsNull(.Fields("PlateNo")), "", .Fields("PlateNo")) & vbTab & IIf(IsNull(.Fields("TypeOfUsage")), "", .Fields("TypeOfUsage")) & vbTab & IIf(IsNull(.Fields("KmReading")), "", .Fields("KmReading"))
            txtRemarks.Text = IIf(IsNull(.Fields("Remarks")), "", .Fields("Remarks"))
            'txtTotal.Text = IIf(IsNull(.Fields("TotalCost")), 0, .Fields("TotalCost"))
            txtPrepby.Text = IIf(IsNull(.Fields("PreparedBy")), "", .Fields("PreparedBy"))
            cboMechanic.Text = IIf(IsNull(.Fields("Mechanic")), "", .Fields("Mechanic"))
            txtOR.Text = IIf(IsNull(.Fields("OrNo")), "", .Fields("OrNo"))
            txtPrintCopy.Text = IIf(IsNull(.Fields("PrintCopy")), 0, .Fields("PrintCopy"))
            txtCustID.Text = IIf(IsNull(!CUSTOMERID), "", !CUSTOMERID)
            txtDesignation.Text = IIf(IsNull(!DESIGNATION), "", !DESIGNATION)
            txtPctD.Text = IIf(IsNull(!DISCOUNT), "", Format(!DISCOUNT, "##0.00") & "%")
            If Not IsNull(!PayType) Then
                If !PayType = "ON ACCOUNT" Then
                    chkOnAcct.Value = 1
                    chkType.Value = 1
                    'chkOnAcct.Visible = True
                Else
                    chkType.Value = 0
                End If
                chkType.Visible = True
            End If
        End If
        .Close
    End With
    '***Access
    'Qry = "SELECT SRVJOBDETAILS.* From SRVJOBDETAILS " & _
          "WHERE (((SRVJOBDETAILS.JOBORDERID)=" & MyNo & "));"
    'mySQl
    Qry = "SELECT SRVJOBDETAILS.* From SRVJOBDETAILS " & _
          "WHERE SRVJOBDETAILS.JOBORDERID=" & MyNo & ";"
    With rsDetails
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxJob.Rows = 1
            .MoveFirst
            Do While Not .EOF
                myDisc = IIf(IsNull(!DISCOUNT), 0, !DISCOUNT)
                MyLabor = IIf(IsNull(!Labor), 0, !Labor)
                myTtl = MyLabor - (MyLabor * (myDisc / 100))
                flxJob.AddItem vbTab & .Fields(1) & vbTab & IIf(IsNull(.Fields(2)), "", .Fields(2)) & vbTab & Format(MyLabor, "#,##0.00") & vbTab & IIf(myDisc <> 0, Format(myDisc, "##0.00") & "%", "") & vbTab & Format(myTtl, "#,##0.00")
                .MoveNext
            Loop
        End If
        .Close
    End With
    ActRow = flxJob.Rows - 1
    Call ComputeTotals
        
    'Display Requisition
    '***Access
    'Qry = "SELECT Requisition.REQNO, Requisition.INVOICENO, Requisition.JOBNO From Requisition " & _
          "WHERE (((Requisition.JOBNO)='" & txtJobNo.Text & "') AND ((Requisition.REMARKS)<>'CANCELLED'));"
    'mySQl
    Qry = "SELECT Requisition.REQNO, Requisition.INVOICENO, Requisition.JOBNO From Requisition " & _
          "WHERE Requisition.JOBNO='" & Trim(txtJobNo.Text) & "' AND Requisition.REMARKS<>'CANCELLED';"
    
    With rsReq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboReq.Clear
            Do While Not .EOF
                cboReq.AddItem !REQNO
                sReq = IIf(sReq <> "", sReq & ", " & !REQNO, "Requisition No. " & !REQNO) & IIf(Not IsNull(!InvoiceNO), " Invoice No. " & !InvoiceNO, "")
                .MoveNext
            Loop
        End If
        .Close
    End With
    'Load Customer list
    Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CONTACTNO, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.TYPEOFUSAGE, SRVCUSTOMER.BRAND, SRVCUSTOMER.ONACCOUNT, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.DESIGNATION " & _
          "From SRVCUSTOMER ORDER BY SRVCUSTOMER.CUSTOMER;"
    If rsCustomer.State = adStateOpen Then rsCustomer.Close
    rsCustomer.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxCust.DataSource = rsCustomer
    With flxCust
        .Width = txtCustomer.Width
        .ColWidth(0) = 0
        .ColWidth(1) = txtCustomer.Width
        .ColWidth(2) = 0    'address
        .ColWidth(3) = 0    'contactno
        .ColWidth(4) = 0    'model
        .ColWidth(5) = 0    'engine
        .ColWidth(6) = 0    'chassis
        .ColWidth(7) = 0    'plateno
        .ColWidth(8) = 0    'Type of Usage
        .ColWidth(9) = 0    'Brand
        .ColWidth(10) = 0   'On Account
        .ColWidth(11) = 0   'CustomerID
        .ColWidth(12) = 0   'Designation
    End With
    If sReq <> "" Then
        flxJob.AddItem vbTab & sReq
        cboReq.Text = cboReq.List(0)
    End If
    cboReq.Locked = False
    cmdReq.Enabled = True
    cmdPO.Enabled = True
End Sub

Function ValidateJobOrder() As String
    Dim vMsg As String, xQry As String
    Dim i As Integer, cntFSC As Integer
    
    vMsg = ""
    If txtCustomer.Text = "" Then vMsg = "Incomplete entry on field Customer Name."
    If txtJobNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "Please enter a valid Job Order No."
    Else
        If AddNew Then
            '***Access
            'xQry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER " & _
                   "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtJobNo.Text & "'));"
            'mySQL
            xQry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER " & _
                   "WHERE SRVCJOBORDER.JOBORDERNO='" & Trim(txtJobNo.Text) & "';"
            
            If CtrlNumExist(xQry) Then
                vMsg = vMsg & vbCrLf & "Job Order No. " & txtJobNo.Text & " already exist!"
            End If
        End If
    End If
    If Not IsDate(txtDateRec.Text) Then vMsg = vMsg & vbCrLf & "Please enter a valid Date Deposit"
    If txtDateProm.Text <> "" Then
        If Not IsDate(txtDateProm.Text) Then vMsg = vMsg & vbCrLf & "Please enter a valid Date Promised"
    End If
    If cboType.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Service Type"
    If UCase$(txtPosMechanic.Text) <> "MACHINIST" Then
        If flxMC.TextMatrix(1, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Type of Usage."
        If flxMC.TextMatrix(1, 4) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field KM Reading."
        If txtAcct.Text = "" Then
            If txtAddress.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Address."
            If flxMC.TextMatrix(1, 0) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Motorcycle Model"
            If flxMC.TextMatrix(1, 1) = "" Then
                vMsg = vMsg & vbCrLf & "Incomplete entry on field Engine Number"
            Else
                '***Access
                'xQry = "SELECT SRVCUSTOMER.ENGINE From SRVCUSTOMER WHERE (((SRVCUSTOMER.ENGINE)='" & flxMC.TextMatrix(1, 1) & "'));"
                'mySQl
                xQry = "SELECT SRVCUSTOMER.ENGINE From SRVCUSTOMER WHERE SRVCUSTOMER.ENGINE='" & flxMC.TextMatrix(1, 1) & "';"
                
                If CtrlNumExist(xQry) Then
                    vMsg = vMsg & vbCrLf & "Engine No. " & flxMC.TextMatrix(1, 1) & " already exist!"
                End If
            End If
            If flxMC.TextMatrix(1, 2) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Chassis Number"
        End If
    End If
    If txtPrepby.Text = "" Then vMsg = vMsg & vbCrLf & "The field Prepared By must not be empty."
    If cboMechanic.Text = "" Then
        vMsg = vMsg & vbCrLf & "The field Mechanic must not be empty"
    Else
        If txtMechanicID.Text = "" Then vMsg = vMsg & vbCrLf & "Please select a Mechanic from the list"
    End If
    If flxJob.TextMatrix(flxJob.Rows - 1, 1) = "" And flxJob.TextMatrix(flxJob.Rows - 1, 3) = "" And flxJob.Rows > 2 Then flxJob.RemoveItem flxJob.Rows - 1
    If flxJob.TextMatrix(flxJob.Rows - 1, 1) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on list of Job Request."
    With flxJob
        cntFSC = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) <> "" Then cntFSC = cntFSC + 1
            If .TextMatrix(i, 2) = "" And Val(.TextMatrix(i, 3)) = 0 Then
               If InStr(.TextMatrix(i, 1), "Requisition") = 0 Then vMsg = vMsg & vbCrLf & "No Labor Amount specified @ row " & i
            End If
        Next i
        If InStr(cboType.Text, "FSC") > 0 And cntFSC = 0 Then vMsg = vMsg & vbCrLf & "No Entry on FSC Number."
    End With
    'If InStr(cboType.Text, "FSC") > 0 Or InStr(cboType.Text, "FREE SERVICE") > 0 Then chkType.Value = 0
    ValidateJobOrder = vMsg
End Function

Function GetSP() As String
    Dim i As Integer
    Dim tmpSP As String
    
    tmpSP = ""
    With flxJob
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 7) <> "" Then
            tmpSP = IIf(tmpSP <> "", tmpSP & "|" & .TextMatrix(i, 7), .TextMatrix(i, 7))
        End If
    Next i
    End With
    GetSP = tmpSP
End Function
Private Sub SaveJobOrder()
    Dim CustSrvc As New ADODB.Recordset
    Dim rsJobOrder As New ADODB.Recordset
    Dim rsJobDetail As New ADODB.Recordset
    Dim i As Integer, OutLabor As Single, Gtotal As Single, myDisc As Single
    Dim errLvl As String, Qry As String
    
    'mySQL Variable
    Dim DatePromise As String
    
    On Error GoTo LocalErr
    If txtAcct.Text = "" Then
        errLvl = "Saving Table Customer Records"
        'With CustSrvc
        '    .Open "SrvCustomer", LYRcon, adOpenKeyset, adLockOptimistic
        '    .AddNew
        '    !CUSTOMER = txtCustomer.Text
        '    !Address = txtAddress.Text
        '    !CONTACTNO = txtContact.Text
        '    !BRAND = cboBrand.Text
        '    !MODEL = flxMC.TextMatrix(1, 0)
        '    !ENGINE = flxMC.TextMatrix(1, 1)
        '    !CHASSIS = flxMC.TextMatrix(1, 2)
        '    !PLATENO = Left$(flxMC.TextMatrix(1, 3), 20)
        '    !TypeOfUsage = flxMC.TextMatrix(1, 4)
        '    !Ctype = "INDIVIDUAL"
        '    If chkAddMC.Value = 1 Then .Fields("ADDTLMC") = 1 'True
        '    .Update
        '    txtAcct.Text = !ACCOUNTNO
        '    .Close
        'End With
        'mySQl
        LYRcon.Execute "INSERT INTO srvcustomer( CUSTOMER, ADDRESS, CONTACTNO, BRAND, MODEL, ENGINE, CHASSIS, PLATENO, TYPEOFUSAGE, CTYPE, ADDTLMC) " & _
            "VALUES('" & txtCustomer.Text & "', '" & txtAddress.Text & "', '" & txtContact.Text & "', '" & cboBrand.Text & "', '" & flxMC.TextMatrix(1, 0) & "', '" & flxMC.TextMatrix(1, 1) & "', " & _
            "'" & flxMC.TextMatrix(1, 2) & "', '" & Left$(flxMC.TextMatrix(1, 3), 20) & "', '" & flxMC.TextMatrix(1, 4) & "', 'INDIVIDUAL', " & chkAddMC.Value & ");"
        Qry = "SELECT srvcustomer.ACCOUNTNO From srvcustomer Where srvcustomer.CUSTOMER = '" & txtCustomer.Text & "' AND srvcustomer.ADDRESS = '" & txtAddress.Text & "' AND srvcustomer.ENGINE = '" & flxMC.TextMatrix(1, 1) & "';"
        txtAcct.Text = GetFieldValue(Qry)
    Else
        '***Access
        'LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.CONTACTNO = '" & txtContact.Text & "', SRVCUSTOMER.BRAND = '" & cboBrand.Text & "', SRVCUSTOMER.MODEL = '" & flxMC.TextMatrix(1, 0) & "', SRVCUSTOMER.PLATENO = '" & flxMC.TextMatrix(1, 3) & "', SRVCUSTOMER.TYPEOFUSAGE = '" & flxMC.TextMatrix(1, 4) & "' " & _
                       "WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & txtAcct.Text & "));"
        'mySQL
        LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.CONTACTNO = '" & txtContact.Text & "', SRVCUSTOMER.BRAND = '" & cboBrand.Text & "', SRVCUSTOMER.MODEL = '" & flxMC.TextMatrix(1, 0) & "', SRVCUSTOMER.PLATENO = '" & flxMC.TextMatrix(1, 3) & "', SRVCUSTOMER.TYPEOFUSAGE = '" & flxMC.TextMatrix(1, 4) & "' " & _
                       "WHERE SRVCUSTOMER.ACCOUNTNO=" & txtAcct.Text & ";"
    End If
    errLvl = "Saving table Job Order"
    'With rsJobOrder
    '    If AddNew Then
    '        .Open "SRVCJOBORDER", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
    '        Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER " & _
    '              "WHERE (((SRVCJOBORDER.JOBORDERID)=" & txtJOid.Text & "));"
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
    '    End If
    '    .Fields("JobOrderNo") = txtJobNo.Text
    '    .Fields("DateReceived") = txtDateRec.Text
    '    If IsDate(txtDateProm.Text) Then
    '        .Fields("DatePromised") = txtDateProm.Text
    '    Else
    '        .Fields("DatePromised") = txtDateRec.Text
    '    End If
    '    .Fields("AccountNo") = txtAcct.Text
    '    OutLabor = Val(Replace(txtTotal.Text, ",", ""))
    '    Gtotal = Val(Replace(txtGrandTtl.Text, ",", ""))
    '    .Fields("TotalCost") = OutLabor
    '    .Fields("Balance") = Gtotal
    '    MyDisc = Val(Replace(txtPctD.Text, "%", ""))
    '    .Fields("Discount") = MyDisc
    '    .Fields("KmReading") = flxMC.TextMatrix(1, 5)
    '    .Fields("Remarks") = txtRemarks.Text
    '    .Fields("Customer") = txtCustomer.Text
    '    .Fields("JOType") = cboType.Text
    '    .Fields("PreparedBy") = txtPrepBy.Text
    '    .Fields("Mechanic") = cboMechanic.Text
    '    .Fields("UpdatedBy") = UserID
    '    .Fields("LastUpdate") = Date
    '    .Fields("EmpID") = txtMechanicID.Text
    '    .Fields("PayType") = chkType.Caption
    '    .Fields("CustomerId") = Val(txtCustID.Text)
    '    .Update
    '    If AddNew Then txtJOid.Text = .Fields("JobOrderID")
    '    .Close
    'End With
    
    
    'mySQL
    OutLabor = Val(Replace(txtTotal.Text, ",", ""))
    Gtotal = Val(Replace(txtGrandTtl.Text, ",", ""))
    myDisc = Val(Replace(txtPctD.Text, "%", ""))
    If IsDate(txtDateProm.Text) Then
        DatePromise = Format(txtDateProm.Text, "yyyy-mm-dd")
    Else
        DatePromise = Format(txtDateRec.Text, "yyyy-mm-dd")
    End If
    If AddNew Then
        LYRcon.Execute "Insert Into srvcjoborder(JOBORDERNO, DATERECEIVED, DATEPROMISED, ACCOUNTNO, TOTALCOST, BALANCE, " & _
            "DISCOUNT, KMREADING, REMARKS, CUSTOMER, JOTYPE, PREPAREDBY, MECHANIC, UPDATEDBY, LASTUPDATE, EMPID, PAYTYPE, CUSTOMERID)" & _
            "VALUES('" & Trim(txtJobNo.Text) & "', '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', '" & DatePromise & "', " & Val(txtAcct.Text) & ", " & _
            "" & OutLabor & ", " & Gtotal & ", " & myDisc & ", '" & flxMC.TextMatrix(1, 5) & "', '" & txtRemarks.Text & "', " & _
            "'" & txtCustomer.Text & "', '" & cboType.Text & "', '" & txtPrepby.Text & "', '" & cboMechanic.Text & "', " & UserID & ", " & _
            "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & txtMechanicID.Text & ", '" & chkType.Caption & "', " & Val(txtCustID.Text) & ");"
        Qry = "select SRVCJOBORDER.JobOrderID from SRVCJOBORDER where SRVCJOBORDER.JOBORDERNO= '" & Trim(txtJobNo.Text) & "';"
        txtJOid.Text = GetFieldValue(Qry)
    Else
        LYRcon.Execute "Update srvcjoborder SET JOBORDERNO = '" & Trim(txtJobNo.Text) & "', DATERECEIVED = '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', " & _
            "DATEPROMISED = '" & DatePromise & "', ACCOUNTNO = " & txtAcct.Text & ", TOTALCOST = " & OutLabor & ", BALANCE = " & Gtotal & ", " & _
            "DISCOUNT = " & myDisc & ", KMREADING = '" & flxMC.TextMatrix(1, 5) & "', REMARKS = '" & txtRemarks.Text & "', CUSTOMER = '" & txtCustomer.Text & "', " & _
            "JOTYPE = '" & cboType.Text & "', PREPAREDBY = '" & txtPrepby.Text & "', MECHANIC = '" & cboMechanic.Text & "', UPDATEDBY = " & UserID & ", " & _
            "LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', EMPID = " & txtMechanicID.Text & ", PAYTYPE = '" & chkType.Caption & "', " & _
            "CUSTOMERID = " & Val(txtCustID.Text) & " Where srvcjoborder.JOBORDERNO = '" & Trim(cusJONO) & "';"
    End If
    If Not AddNew Then
        errLvl = "Removing old Job Order Details"
        'Access
        'LYRcon.Execute "DELETE SRVJOBDETAILS.*, SRVJOBDETAILS.JOBORDERID " & _
                       "From SRVJOBDETAILS WHERE (((SRVJOBDETAILS.JOBORDERID)=" & txtJOid.Text & "));"
        'mySQL
        LYRcon.Execute "DELETE From SRVJOBDETAILS WHERE SRVJOBDETAILS.JOBORDERID=" & txtJOid.Text & ";"
    End If
    errLvl = "Saving Job Order Details"
    '***Access
    'With rsJobDetail
        '.Open "SRVJobDetails", LYRcon, adOpenKeyset, adLockOptimistic
        For i = 1 To flxJob.Rows - 1
            If InStr(UCase$(flxJob.TextMatrix(i, 1)), "REQUISITION") = 0 Then
                '***Access
                '.AddNew
                '.Fields("JobOrderId") = txtJOid.Text
                '.Fields("JobRequest") = flxJob.TextMatrix(i, 1)
                '.Fields("FSC") = flxJob.TextMatrix(i, 2)
                '.Fields("Labor") = PureNum(flxJob.TextMatrix(i, 3))
                '.Fields("Discount") = Val(Replace(flxJob.TextMatrix(i, 4), "%", ""))
                '.Fields("Total") = PureNum(flxJob.TextMatrix(i, 5))
                '.Update
                'mySQL
                LYRcon.Execute "INSERT INTO srvjobdetails( JOBORDERID, JOBREQUEST, FSC, LABOR, DISCOUNT, TOTAL)" & _
                    "VALUES(" & Val(txtJOid.Text) & ", '" & flxJob.TextMatrix(i, 1) & "', '" & flxJob.TextMatrix(i, 2) & "', " & _
                    "" & Val(Replace(flxJob.TextMatrix(i, 3), ",", "")) & ", " & Val(Replace(flxJob.TextMatrix(i, 4), "%", "")) & ", " & _
                    "" & Val(Replace(flxJob.TextMatrix(i, 5), ",", "")) & ");"
            End If
        Next i
        '.Close
    'End With
    Exit Sub
LocalErr:
    MsgBox errLvl & vbCrLf & str(Err.Number) & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Function StartPay(S As String) As Date
    Dim iPS As Integer, sPay As Date, MyDate As Date
    
    MyDate = IIf(IsDate(txtDateProm.Text), txtDateProm.Text, txtDateRec.Text)
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

Private Sub UpdateJOBonAcct()
    Dim rsOnAccount As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim nPay As Integer, nTrm As Integer, Bal As Currency
    Dim Qry As String, nSkm As String
    
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
            nTrm = !Terms
            nSkm = !PayScheme
            nPay = !Npays
        Else
            nTrm = 30
            nSkm = "MONTHLY"
            nPay = 1
        End If
        .Close
    End With
        Gtotal = Val(Replace(txtGrandTtl.Text, ",", ""))
    '***Access
    'With rsOnAccount
    '    If AddNew Then
    '        .Open "ONACCTERMS", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
    '        Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
    '              "WHERE (((ONACCTERMS.JOBORDERNO)='" & cmdTrans(6).Tag & "') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & ") AND ((ONACCTERMS.SPINVNO) Is Null));"
    '        If Not CtrlNumExist(Qry) Then
    '            .Open "ONACCTERMS", LYRcon, adOpenKeyset, adLockOptimistic
    '            .AddNew
    '        Else
    '            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '            .MoveFirst
    '        End If
    '    End If
        
    '    !CUSTOMERID = txtCustID.Text
    '    !PayScheme = nSkm
    '    !Terms = nTrm
    '    !Npays = nPay
    '    txtPayStart.Text = StartPay(nSkm)
    '    !PayStart = DateValue(txtPayStart)
    '    !BILLINGAMT = Format(Gtotal / nPay, "###0.00")
    '    !JOBORDERNO = txtJobNo.Text
    '    !JOBORDERID = Val(txtJOid.Text)
    '    !BALANCE = Gtotal
    '    !AcctAmount = Gtotal
    '    !TRANSDATE = Format$(Date, "mm/dd/yyyy")
    '    !Status = "OPEN"
    '    !LASTUPDATE = Date
        'old Amout Due, Due Date , last Update
    '    .Update
    '    .Close
        
        
        'mySQL Statement
        txtPayStart.Text = StartPay(nSkm)
        If AddNew Then
                LYRcon.Execute "Insert Into onaccterms(CUSTOMERID, PAYSCHEME, TERMS, NPAYS, PAYSTART, " & _
                    "BILLINGAMT, JOBORDERNO, JOBORDERID, BALANCE, ACCTAMOUNT, TRANSDATE, `STATUS`, " & _
                    "LASTUPDATE)VALUES(" & txtCustID.Text & ", '" & nSkm & "', " & nTrm & ", " & nPay & ", " & _
                    "'" & Format(DateValue(txtPayStart), "yyyy-mm-dd") & "', " & Format(Gtotal / nPay, "###0.00") & ", " & _
                    "'" & Trim(txtJobNo.Text) & "', " & Val(txtJOid.Text) & ", " & Gtotal & ", " & Gtotal & ", " & _
                    "'" & Format(Date, "yyyy-mm-dd") & "', 'OPEN', '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        Else
            Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
                "WHERE ONACCTERMS.JOBORDERNO='" & cmdTrans(6).Tag & "' AND ONACCTERMS.CUSTOMERID=" & txtCustID.Text & " AND (ONACCTERMS.SPINVNO) Is Null;"
            If Not CtrlNumExist(Qry) Then
                LYRcon.Execute "Insert Into onaccterms(CUSTOMERID, PAYSCHEME, TERMS, NPAYS, PAYSTART, " & _
                    "BILLINGAMT, JOBORDERNO, JOBORDERID, BALANCE, ACCTAMOUNT, TRANSDATE, `STATUS`, " & _
                    "LASTUPDATE)VALUES(" & txtCustID.Text & ", '" & nSkm & "', " & nTrm & ", " & nPay & ", " & _
                    "'" & Format(DateValue(txtPayStart), "yyyy-mm-dd") & "', " & Format(Gtotal / nPay, "###0.00") & ", " & _
                    "'" & Trim(txtJobNo.Text) & "', " & Val(txtJOid.Text) & ", " & Gtotal & ", " & Gtotal & ", " & _
                    "'" & Format(Date, "yyyy-mm-dd") & "', 'OPEN', '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
           Else
                LYRcon.Execute "Update onaccterms SET CUSTOMERID = " & txtCustID.Text & ", PAYSCHEME = '" & nSkm & "', " & _
                    "TERMS = " & nTrm & ", NPAYS = " & nPay & ", PAYSTART = '" & Format(DateValue(txtPayStart), "yyyy-mm-dd") & "', " & _
                    "BILLINGAMT = " & Format(Gtotal / nPay, "###0.00") & ", JOBORDERNO = '" & Trim(txtJobNo.Text) & "', JOBORDERID = " & Val(txtJOid.Text) & ", " & _
                    "BALANCE = " & Gtotal & ", ACCTAMOUNT = " & Gtotal & ", TRANSDATE = '" & Format(Date, "yyyy-mm-dd") & "', `STATUS` = 'OPEN', " & _
                    "LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                    "WHERE ONACCTERMS.JOBORDERNO='" & cmdTrans(6).Tag & "' AND ONACCTERMS.CUSTOMERID=" & txtCustID.Text & " AND (ONACCTERMS.SPINVNO) Is Null;"
           End If
        End If
    'End With
    'Update Customer Outstanding Balance
    '***Access
    'Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS WHERE (((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
    'mySQL
    Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS WHERE ONACCTERMS.CUSTOMERID=" & txtCustID.Text & " AND ONACCTERMS.STATUS='OPEN';"
    
    Bal = Val(GetFieldValue(Qry))
    '***Access
    'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & Bal & _
                   " WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & Bal & _
                   " WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
End Sub
Private Sub ClearEntries()
    txtCustomer.Text = ""
    txtAddress.Text = ""
    txtContact.Text = ""
    txtAcct.Text = ""
    cboType.Text = ""
    txtJobNo.Text = ""
    txtDateRec.Text = ""
    txtDateProm.Text = ""
    txtDown.Text = ""
    txtBal.Text = ""
    flxMC.Rows = 1
    flxMC.Rows = 2
    flxJob.Rows = 1
    flxJob.Rows = 2
    txtRemarks.Text = ""
    txtTotal.Text = ""
    cboMechanic.Text = ""
    cboReq.Clear
    'chkOnAcct.Visible = True
    chkOnAcct.Value = 1
    chkType.Value = 0
    txtCustID.Text = ""
    txtPayStart.Text = ""
    txtDiscAmt.Text = ""
    txtPctD.Text = ""
    txtGrandTtl.Text = ""
    txtOR.Text = ""
    txtPrintCopy.Text = ""
    lblAvlMC.Caption = ""
    cboEngine.Clear
End Sub

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    flxMC.Enabled = Not bln
    flxJob.Enabled = Not bln
    'Frame2.Enabled = Not Bln
    txtOR.Locked = bln
    txtPrepby.Locked = bln
    cboMechanic.Locked = bln
    cboReq.Locked = Not bln
    cmdReq.Enabled = bln
    cmdPO.Enabled = bln
    txtRemarks.Locked = bln
    txtPctD.Locked = bln
    txtDiscAmt.Locked = bln
End Sub

Private Sub cboBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboEngine_Change()
   Call cboEngine_Click
End Sub

Private Sub cboEngine_Click()
    Dim i As Integer, xRow As Integer
    
    With flxAllMC
        xRow = 0
        For i = 1 To .Rows - 1
            If cboEngine.Text = .TextMatrix(i, 5) Then
                xRow = i
                Exit For
            End If
        Next i
        If xRow = 0 Then
            txtAddress.Text = ""
            txtAcct.Text = ""
            cboBrand.Text = ""
            txtCustID.Text = "" 'CustomerID
            txtContact.Text = ""
            flxMC.TextMatrix(1, 0) = "" 'Model
            'flxMC.TextMatrix(1, 1) = "" 'Engine
            flxMC.TextMatrix(1, 2) = "" 'chassis
            flxMC.TextMatrix(1, 3) = "" 'Plate
            flxMC.TextMatrix(1, 4) = "" 'type of Usage
            chkOnAcct.Value = 0 'On Acct
            chkType.Value = 0
            txtDesignation.Text = ""
            Exit Sub
        End If
        'Insert Additional retrieval here
        txtAddress.Text = .TextMatrix(xRow, 2)
        txtAcct.Text = .TextMatrix(xRow, 0)
        cboBrand.Text = .TextMatrix(xRow, 9) 'Brand
        txtCustID.Text = .TextMatrix(xRow, 11)    'CustomerID
        txtContact.Text = .TextMatrix(xRow, 3)
        flxMC.TextMatrix(1, 0) = .TextMatrix(xRow, 4) 'Model
        'flxMC.TextMatrix(1, 1) = .TextMatrix(xRow, 5) 'Engine
        flxMC.TextMatrix(1, 2) = .TextMatrix(xRow, 6) 'chassis
        flxMC.TextMatrix(1, 3) = .TextMatrix(xRow, 7) 'Plate
        flxMC.TextMatrix(1, 4) = .TextMatrix(xRow, 8) 'type of Usage
        If .TextMatrix(xRow, 10) = "1" Then
            chkOnAcct.Value = 1 'On Acct
            'chkType.Caption = "ON-ACCOUNT"
            'chkType.ForeColor = vbRed
            chkType.Value = 1 'On Account Type
            txtDesignation.Text = .TextMatrix(xRow, 12) 'Designation
        Else
            chkOnAcct.Value = 0 'On Acct
            'chkType.Caption = "CASH"
            'chkType.ForeColor = vbGree
            chkType.Value = 0
            txtDesignation.Text = ""
        End If
    
    End With
End Sub

Private Sub cboEngine_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxMC.TextMatrix(MCrow, MCcol) = cboEngine.Tag
        cboEngine.Tag = "skiplostfocus"
        cboEngine.Visible = False
        flxMC.SetFocus
        flxMC.Col = 5
    End If
End Sub

Private Sub cboEngine_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxMC.SetFocus
        flxMC.Col = 2
    End If
End Sub

Private Sub cboEngine_LostFocus()
    If cboEngine.Tag = "skiplostfocus" Then
        cboEngine.Tag = ""
        Exit Sub
    End If
    If cboEngine.ListIndex = -1 Then
        MsgBox "Please select an Engine from the list!", vbExclamation + vbOKOnly, "Invalid Engine"
        cboEngine.Visible = False
        Exit Sub
    End If
    flxMC.TextMatrix(MCrow, 1) = cboEngine.Text
    cboEngine.Visible = False
End Sub

Private Sub cboJob_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If JobSrch Then
        JobSrch = False
        Exit Sub
    End If
    With cboJob
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsJobDesc.MoveFirst
            rsJobDesc.Move i
            If flxJob.TextMatrix(ActRow, 2) = "" Or UCase$(cboType.Text) = "QUICK SERVICE" Then
                flxJob.TextMatrix(ActRow, 3) = IIf(IsNull(rsJobDesc!LaborCost), "", Format$(rsJobDesc!LaborCost, "###,##0.00"))
            End If
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                JobSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            flxJob.TextMatrix(ActRow, 3) = ""
        End If
    End With
    'old flxJob.TextMatrix(ActRow, 6) = PureNum(flxJob.TextMatrix(ActRow, 3)) + PureNum(flxJob.TextMatrix(ActRow, 4)) + PureNum(flxJob.TextMatrix(ActRow, 5))
    Call ComputeTotals
End Sub

Private Sub cboJob_Click()
    Call cboJob_Change
End Sub

Private Sub cboJob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then JobSrch = True
    If KeyCode = vbKeyEscape Then
        flxJob.TextMatrix(ActRow, ActCol) = cboJob.Tag
        cboJob.Tag = "skiplostfocus"
        cboJob.Visible = False
        flxJob.SetFocus
        flxJob.Col = 1
    End If
End Sub

Private Sub cboJob_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxJob.SetFocus
        flxJob.Col = 2
    End If
End Sub

Private Sub cboJob_LostFocus()
    If cboJob.Tag = "skiplostfocus" Then
        cboJob.Tag = ""
        Exit Sub
    End If
    flxJob.TextMatrix(ActRow, ActCol) = cboJob.Text
    cboJob.Visible = False
    Call ComputeTotals
End Sub


Private Sub cboMechanic_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If MechSrch Then
        MechSrch = False
        Exit Sub
    End If
    With cboMechanic
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsMechanic.MoveFirst
            rsMechanic.Move i
            txtMechanicID.Text = rsMechanic!EMPID
            txtPosMechanic.Text = rsMechanic!Position
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                MechSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtMechanicID.Text = ""
            txtPosMechanic.Text = ""
        End If
    End With
End Sub

Private Sub cboMechanic_Click()
    Call cboMechanic_Change
End Sub

Private Sub cboMechanic_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboMechanic.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cboMechanic_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            MechSrch = True
            If rsMechanic.State = adStateOpen Then rsMechanic.MoveFirst
        Case vbKeyBack
            MechSrch = True
            If rsMechanic.State = adStateOpen Then rsMechanic.MoveFirst
    End Select

End Sub


Private Sub cboType_GotFocus()
    txtSrvcType.Text = cboType.Text
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJobNo.SetFocus
End Sub

Private Sub cboType_LostFocus()
    If txtSrvcType.Text <> "" Then
        txtRemarks.Text = Replace(txtRemarks.Text, txtSrvcType, "")
    End If
    If cboType.Text <> "QUICK SERVICE" Then
        txtRemarks.Text = cboType.Text & " " & txtRemarks.Text
        If cboType.Text = "FREE SERVICE" Then
            txtPctD.Text = "100.00%"
            txtDiscAmt.Text = txtTotal.Text
            txtGrandTtl.Text = "0.00"
        End If
        
    End If
End Sub

Private Sub chkType_Click()
    If chkType.Value = 1 Then
        'Check if customer is accredited for On-Account
        If chkOnAcct.Value = 0 Then
            MsgBox "The customer is not accredited for On-Account!" & vbCrLf & _
                   "Please update customer status to the manager.", vbExclamation + vbOKOnly, "Not Accredited"
            chkType.Value = 0
            Exit Sub
        End If
        chkType.Caption = "ON ACCOUNT"
        chkType.ForeColor = vbRed
    Else
        chkType.Caption = "CASH"
        chkType.ForeColor = vbGreen
    End If
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdNewMC_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Add a different motorcycle for " & txtCustomer.Text & "?", vbQuestion + vbYesNo, "Another MC")
    If Ans = vbNo Then Exit Sub
    For i = 0 To 4
        flxMC.TextMatrix(1, i) = ""
    Next i
    
    'If address is company then clear address
    If Val(txtAcct.Text) > 0 Then
        If CustIsComp(txtAcct.Text) Then txtAddress.Text = ""
    End If
    txtAcct.Text = ""
    cboBrand.Text = ""
    txtCustID.Text = "" 'CustomerID
    txtContact.Text = ""
    'Additional MC is always Cash
    chkOnAcct.Value = 0
    chkType.Value = 0
    chkAddMC.Value = 1
    cmdNewMC.Visible = False
    cboBrand.SetFocus
End Sub
Function CustIsComp(MyID As String) As Boolean
    Dim rsCust As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SRVCUSTOMER.CTYPE From SRVCUSTOMER WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & MyId & "));"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.CTYPE From SRVCUSTOMER WHERE SRVCUSTOMER.ACCOUNTNO=" & MyID & ";"
    With rsCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            If Not IsNull(!Ctype) Then
                If !Ctype = "COMPANY" Then
                    CustIsComp = True
                Else
                    CustIsComp = False
                End If
            Else
                CustIsComp = False
            End If
        Else
            CustIsComp = False
        End If
        .Close
    End With
End Function

Private Sub cmdPO_Click()
    If chkType.Value = 0 Then
        MsgBox "P.O. is for On-Account type only!", vbExclamation + vbOKOnly, "Job Order"
        Exit Sub
    Else
        POCust.MySource = "JOB ORDER"
        POCust.Show
    End If
End Sub

Private Sub cmdReq_Click()
    Requisition.MySource = "Job Order"
    Requisition.Show vbModal
End Sub

Function GetAcctType(AcctID As Long) As Integer
    Dim rsAcct As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.ONACCOUNT " & _
          "From SRVCUSTOMER WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & AcctID & "));"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.ONACCOUNT " & _
          "From SRVCUSTOMER WHERE SRVCUSTOMER.ACCOUNTNO=" & AcctID & ";"
    
    With rsAcct
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            'Access
            GetAcctType = IIf(!ONACCOUNT = True, 1, 0)
            'mySQL
            'GetAcctType = IIf(!ONACCOUNT = 1, 1, 0)
        Else
            GetAcctType = 0
        End If
        .Close
    End With
    
End Function

Function ChargeToAcct(AcctID As Variant) As String
    Dim rsAcct As New ADODB.Recordset
    Dim Qry As String, MyType As String
    
    '***Access
    'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CTYPE From SRVCUSTOMER WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & AcctID & "));"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CTYPE From SRVCUSTOMER WHERE SRVCUSTOMER.ACCOUNTNO=" & AcctID & ";"
    
    With rsAcct
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyType = IIf(IsNull(!Ctype), "", !Ctype)
            If MyType = "COMPANY" Then
                ChargeToAcct = txtAddress.Text
            Else
                ChargeToAcct = txtCustomer.Text
            End If
        Else
            ChargeToAcct = ""
        End If
        .Close
    End With
End Function


Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, pdJO As Single, MyTtlJO As Single
    Dim myMsg As String, PrnVal(0 To 26) As String, Qry As String, MyStat As String
    Dim MyAcc As String
    
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=38));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=38;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            BtnIndx = 5
            tmrFloat.Enabled = True
            Call ClearEntries
            '***Access
            'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CONTACTNO, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.TYPEOFUSAGE, SRVCUSTOMER.BRAND, SRVCUSTOMER.ONACCOUNT, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.ADDTLMC " & _
                  "From SRVCUSTOMER Where (((SRVCUSTOMER.ADDTLMC) = False)) ORDER BY SRVCUSTOMER.CUSTOMER;"
            'mySQL
            Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CONTACTNO, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.TYPEOFUSAGE, SRVCUSTOMER.BRAND, SRVCUSTOMER.ONACCOUNT, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.ADDTLMC " & _
                  "From SRVCUSTOMER Where SRVCUSTOMER.ADDTLMC = 0 ORDER BY SRVCUSTOMER.CUSTOMER;"
            
            If rsCustomer.State = adStateOpen Then rsCustomer.Close
            rsCustomer.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            Set flxCust.DataSource = rsCustomer
            With flxCust
                .Width = txtCustomer.Width
                .ColWidth(0) = 0
                .ColWidth(1) = txtCustomer.Width
                .ColWidth(2) = 0    'address
                .ColWidth(3) = 0    'contactno
                .ColWidth(4) = 0    'model
                .ColWidth(5) = 0    'engine
                .ColWidth(6) = 0    'chassis
                .ColWidth(7) = 0    'plateno
                .ColWidth(8) = 0    'Type of Usage
                .ColWidth(9) = 0    'Brand
                .ColWidth(10) = 0   'On Account
                .ColWidth(11) = 0   'CustomerID
                .ColWidth(12) = 0   'Designation
                .ColWidth(13) = 0   'Designation
            End With
            Call LockData(False)
            txtJobNo.Text = GetJobOrderNo
            txtCustomer.Locked = False
            txtJobNo.Locked = False
            txtDateRec.Locked = False
            txtDateRec.Text = Date
            txtPrepby.Text = GetUserName
            txtCustomer.SetFocus
            chkType.Value = 0
            'cboUsage.Text = "Solo"
            'txtPreparedBy.Text = GetSettings("STOCK CLERK")
            AddNew = True
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=38));"
            'mySQl
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=38;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtJobNo.Text = "" Then
                MsgBox "Unable to update empty Job Order!", vbExclamation + vbOKOnly, "Empty"
                Exit Sub
            End If
            If Val(txtPrintCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because Job Order " & txtJobNo.Text & " was already printed.", vbInformation + vbOKOnly
                Exit Sub
            End If
            MyTtlJO = PureNum(txtTotal.Text)
            If MyTtlJO > 0 Then
                pdJO = AmtPaidJO
                If pdJO = MyTtlJO Then
                    MsgBox "Payment on Out Labor already made!", vbExclamation + vbOKOnly, "Editing Failed"
                    Exit Sub
                End If
            End If
            chkOnAcct.Value = GetAcctType(Val(txtAcct.Text))
            BtnIndx = 5
            tmrFloat.Enabled = True
            Call LockData(False)
            cmdTrans(6).Tag = txtJobNo.Text 'Store old JO# to Save Button for reference in updating On-Account
            AddNew = False
            cusJONO = txtJobNo.Text
        Case 2  'View
            'JOSummary.Show
            FrmSummary.mySrc = "Job Order"
            FrmSummary.Show
            Unload Me
        Case 3  'Service Record
            JOBonacct.Show vbModal
        Case 4  'Print
            If txtCustomer.Text = "" Then
                MsgBox "No data available for printing", vbExclamation, "PRINT"
                Exit Sub
            End If
            MyTtlJO = Val(Replace(txtTotal.Text, ",", ""))
            Ans = MsgBox("Print Job Order?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            If InStr(SysBranch, "LYR MOTOR SHOP") > 0 Or SysBranch = "BINOS SPAREPARTS STO. TOMAS" Or SysBranch = "BINOS SPAREPARTS SAMAL" Or SysBranch = "MOTORHUB SPAREPARTS MARBEL" Or SysBranch = "BINOS SPAREPARTS POLOMOLOK" Or SysBranch = "MOTORHUB SPAREPARTS TAGUM" Or SysBranch = "YAMAHA MOTORHUB ULAS" Or SysBranch = "MRC PIAGGIO TAGUM" Then
            'Or SysBranchCategory = "3S" Then  'Or InStr(SysBranch, "PANABO") > 0
            'If InStr(SysBranch, "LYR MOTOR SHOP") > 0 Then
                
                If Val(txtPrintCopy.Text) = 0 Then
                    If chkType.Value = 1 And Val(txtCustID.Text) > 0 And Val(Replace(txtGrandTtl.Text, ",", "")) > 0 Then Call UpdateJOBonAcct
                    '***Access
                    'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.DATEPROMISED = #" & Date & "# " & _
                                   "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtJobNo.Text & "'));"
                    'mySQL
                    LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.DATEPROMISED = '" & Format(Date, "yyyy-mm-dd") & "' " & _
                                   "WHERE SRVCJOBORDER.JOBORDERNO='" & Trim(txtJobNo.Text) & "';"
                    
                End If
                PRNatt.MySource = "JOB ORDER"
                PRNatt.Show vbModal
            Else
                PrnVal(0) = txtCustomer.Text
                PrnVal(1) = txtAddress.Text
                PrnVal(2) = txtContact.Text
                PrnVal(3) = txtDateRec.Text
                PrnVal(4) = Format$(Date, "MM/DD/YYYY") & IIf(DateValue(txtDateRec.Text) <> Date, " [OUT]", "") ' txtDateProm.Text
                PrnVal(5) = IIf(Val(txtDown.Text) > 0, txtDown.Text, "")
                PrnVal(6) = flxMC.TextMatrix(1, 0)  'Model
                PrnVal(7) = flxMC.TextMatrix(1, 3)  'Plate
                PrnVal(8) = flxMC.TextMatrix(1, 4)  'Usage
                PrnVal(9) = flxMC.TextMatrix(1, 5)  'KM Reading
                PrnVal(10) = IIf(Val(txtBal.Text) > 0, txtBal.Text, "")
                PrnVal(11) = flxMC.TextMatrix(1, 2)  'Chassis
                PrnVal(12) = flxMC.TextMatrix(1, 1)  'Engine
                PrnVal(13) = "J O B  D E S C R I P T I O N"
                PrnVal(14) = "FSC#"
                PrnVal(15) = "LABOR"
                PrnVal(16) = "DISC."
                PrnVal(17) = "T O T A L"
                PrnVal(18) = IIf(cboMechanic.Text <> "", UCase$(txtPosMechanic.Text) & ": " & cboMechanic.Text, "")
                PrnVal(19) = IIf(txtRemarks.Text <> "", "NOTE: " & txtRemarks.Text, "") & IIf(chkType.Value = 1, "  ON ACCOUNT: " & ChargeToAcct(Val(txtAcct.Text)), "")
                'PrnVal(19) = IIf(txtRemarks.Text <> "", "NOTE: " & txtRemarks.Text, "")
                PrnVal(20) = "TOTAL:"
                PrnVal(21) = txtTotal.Text
                PrnVal(22) = IIf(txtOR.Text <> "", "O.R. NUMBER: " & txtOR.Text, "")
                If Val(txtDiscAmt.Text) > 0 Then
                    PrnVal(23) = "Less " & txtPctD.Text & " DISCOUNT:"
                    PrnVal(24) = txtDiscAmt.Text
                    PrnVal(25) = "GRAND TOTAL:"
                    PrnVal(26) = txtGrandTtl.Text
                Else
                    PrnVal(23) = ""
                    PrnVal(24) = ""
                    PrnVal(25) = ""
                    PrnVal(26) = ""
                End If
                Call PrintMyForm(PrnVal, "JOB ORDER", flxJob)
                If Val(txtPrintCopy.Text) = 0 Then
                    If chkType.Value = 1 And Val(txtCustID.Text) > 0 And Val(Replace(txtGrandTtl.Text, ",", "")) > 0 Then Call UpdateJOBonAcct
                    '***Access
                    'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.DATEPROMISED = #" & Date & "# " & _
                                   "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtJobNo.Text & "'));"
                    'mySQL
                    LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.DATEPROMISED = '" & Format(Date, "yyyy-mm-dd") & "' " & _
                                   "WHERE SRVCJOBORDER.JOBORDERNO='" & Trim(txtJobNo.Text) & "';"
                    
                End If
                txtPrintCopy.Text = Val(txtPrintCopy.Text) + 1
                '***Access
                'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.PRINTCOPY = " & Val(txtPrintCopy.Text) & _
                               " WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtJobNo.Text & "'));"
                'mySQl
                LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.PRINTCOPY = " & Val(txtPrintCopy.Text) & _
                               " WHERE SRVCJOBORDER.JOBORDERNO ='" & Trim(txtJobNo.Text) & "';"
            End If
        Case 5  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
        Case 6  'Save
            flxCust.Visible = False
            myMsg = ValidateJobOrder()
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Job Order?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 7
            tmrFloatRev.Enabled = True
            Call SaveJobOrder
            Call LockData(True)
            cmdTrans(6).Tag = ""
            MsgBox "Job Order successfully saved!", vbInformation + vbOKOnly, "Save"
        
        Case 7  'Cancel
            flxCust.Visible = False
            Ans = MsgBox("Do you want to Cancel Job Order?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 7
            tmrFloatRev.Enabled = True
            Call ClearEntries
            If Not AddNew Then
                Call DisplayJobOrder(cmdTrans(6).Tag)
                txtCustomer.Locked = False
                txtJobNo.Locked = False
                txtDateRec.Locked = False
                cmdTrans(6).Tag = ""
            End If
            Call LockData(True)
    End Select
End Sub

Function GetJobOrderNo() As String
    Dim rsInvNo As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    '***Access
    'Qry = "SELECT Max(Val([JOBORDERNO])) AS MaxJO FROM SRVCJOBORDER;"
    'mySQl
'    Qry = "SELECT Max(ROUND(SRVCJOBORDER.JOBORDERNO)) AS MaxJO FROM SRVCJOBORDER;"
    Qry = "SELECT Max(Round(SRVCJOBORDER.JOBORDERNO)) As MaxJO From SRVCJOBORDER Where SRVCJOBORDER.DATERECEIVED = '" & Format(Date, "yyyy-mm-dd") & "';"
    
    
    '
    With rsInvNo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 0, .Fields(0) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    GetJobOrderNo = Format(MyNo, "00000")
End Function

Function AmtPaidJO() As Single
    Dim rsJOpay As New ADODB.Recordset
    Dim Qry As String

    With rsJOpay
        '***Access
        'Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
              "WHERE (((PAYMENTS.REMARKS) Like '%" & txtJobNo.Text & "%')) HAVING (((Sum(PAYMENTS.SERVICE))>0));"
        'mySQL
        'Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
              "WHERE PAYMENTS.REMARKS Like '%" & txtJobNo.Text & "%' HAVING Sum(PAYMENTS.SERVICE)>0;"
        Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
              "WHERE PAYMENTS.REMARKS Like '%" & txtJobNo.Text & "%' AND PAYMENTS.REMARKS Like '%JO#%' HAVING Sum(PAYMENTS.SERVICE)>0;"
              
    
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            AmtPaidJO = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            AmtPaidJO = 0
        End If
        .Close
    End With
    
End Function

Private Sub flxCust_Click()
    txtCustomer.Text = flxCust.TextMatrix(flxCust.Row, 1)
End Sub

Private Sub flxCust_EnterCell()
    CustSrch = True
    txtCustomer.Text = flxCust.TextMatrix(flxCust.Row, 1)
End Sub

Private Sub flxCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxCust.Visible = False
        txtCustomer.SetFocus
    End If
End Sub

Private Sub flxCust_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtCustomer_Change
        Call txtCustomer_LostFocus
        flxCust.Visible = False
        txtAddress.SetFocus
    End If
End Sub

Private Sub flxJob_Click()
    ActRow = flxJob.Row
    ActCol = flxJob.Col
    
    If ActCol = 0 Or ActRow = 0 Or ActCol = 6 Then Exit Sub
    Select Case ActCol
    Case 1
        With cboJob
            If flxJob.TextMatrix(ActRow, ActCol) <> "" Then .Tag = flxJob.TextMatrix(ActRow, ActCol)
            .Top = flxJob.CellTop + flxJob.Top
            .Left = flxJob.CellLeft + flxJob.Left
            .Width = flxJob.CellWidth
            .Text = flxJob.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 2
        With txtFSC
            If flxJob.TextMatrix(ActRow, 2) <> "" Then .Tag = flxJob.TextMatrix(ActRow, 2)
            .Top = flxJob.CellTop + flxJob.Top
            .Left = flxJob.CellLeft + flxJob.Left
            .Width = flxJob.CellWidth
            .Text = flxJob.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 3
        With txtFlx
            If flxJob.TextMatrix(ActRow, 3) <> "" Then .Tag = flxJob.TextMatrix(ActRow, 3)
            .Top = flxJob.CellTop + flxJob.Top
            .Left = flxJob.CellLeft + flxJob.Left
            .Width = flxJob.CellWidth
            .Text = flxJob.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 4
        With txtDiscount
            If flxJob.TextMatrix(ActRow, 4) <> "" Then .Tag = flxJob.TextMatrix(ActRow, 4)
            .Top = flxJob.CellTop + flxJob.Top
            .Left = flxJob.CellLeft + flxJob.Left
            .Width = flxJob.CellWidth
            .Text = Replace(flxJob.Text, "%", "")
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    
    End Select
End Sub

Private Sub flxJob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxJob_Click
End Sub

Private Sub flxJob_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuFlx)
    End If
End Sub

Private Sub flxMC_Click()
    Dim ImOut As Boolean
    MCrow = flxMC.Row
    MCcol = flxMC.Col
    
    ImOut = False
    If MCrow = 0 Then ImOut = True
    If MCcol < 4 And txtAcct.Text <> "" Then ImOut = True
    If chkAddMC.Value = 1 Or (MCcol = 1 And cboEngine.ListCount > 1) Then ImOut = False
    If ImOut Then Exit Sub
    Select Case MCcol
        Case 1
            If cboEngine.ListCount > 1 And chkAddMC.Value = 0 Then
                With cboEngine
                    If flxMC.TextMatrix(MCrow, 1) <> "" Then .Tag = flxMC.TextMatrix(MCrow, 1)
                    .Top = flxMC.CellTop + flxMC.Top
                    .Left = flxMC.CellLeft + flxMC.Left
                    .Width = flxMC.CellWidth
                    .Text = flxMC.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Else
                With txtMC
                    If flxMC.TextMatrix(MCrow, 1) <> "" Then .Tag = flxMC.TextMatrix(MCrow, 1)
                    .Top = flxMC.CellTop + flxMC.Top
                    .Left = flxMC.CellLeft + flxMC.Left
                    .Width = flxMC.CellWidth
                    .Text = flxMC.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            End If
        Case 0, 2, 3, 5
            With txtMC
                If flxMC.TextMatrix(MCrow, MCcol) <> "" Then .Tag = flxMC.TextMatrix(MCrow, MCcol)
                .Top = flxMC.CellTop + flxMC.Top
                .Left = flxMC.CellLeft + flxMC.Left
                .Width = flxMC.CellWidth
                .Text = flxMC.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 4
            With cboUsage
                If flxMC.TextMatrix(MCrow, MCcol) <> "" Then .Tag = flxMC.TextMatrix(MCrow, MCcol)
                .Top = flxMC.CellTop + flxMC.Top
                .Left = flxMC.CellLeft + flxMC.Left
                .Width = flxMC.CellWidth
                .Text = flxMC.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
    End Select

End Sub

Private Sub flxMC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxMC_Click
End Sub

Private Sub Form_Load()
    
    Call LockData(True)
    Call LoadCustomer
    
    cboBrand.AddItem "HONDA"
    cboBrand.AddItem "KAWASAKI"
    cboBrand.AddItem "SUZUKI"
    cboBrand.AddItem "YAMAHA"
    With cboType
        .AddItem "BACK JOB"
        .AddItem "DUE TO REBUILT"
        .AddItem "FREE SERVICE"
        .AddItem "FSC #1 HONDA"
        .AddItem "FSC #2 HONDA"
        .AddItem "FSC #3 HONDA"
        .AddItem "FSC #1 KAWASAKI"
        .AddItem "FSC #2 KAWASAKI"
        .AddItem "FSC #3 KAWASAKI"
        .AddItem "FSC #4 KAWASAKI"
        .AddItem "FSC #1 SUZUKI"
        .AddItem "FSC #2 SUZUKI"
        .AddItem "FSC #3 SUZUKI"
        .AddItem "FSC #4 SUZUKI"
        .AddItem "FSC #1 YAMAHA"
        .AddItem "FSC #2 YAMAHA"
        .AddItem "FSC #3 YAMAHA"
        .AddItem "FSC #4 YAMAHA"
        .AddItem "FSC #5 YAMAHA"
        .AddItem "FSC #6 YAMAHA"
        .AddItem "MACHINE SHOP"
        .AddItem "MAJOR SERVICE"
        .AddItem "QUICK SERVICE"
        .AddItem "WARRANTY SERVICE"
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""

    If rsCustomer.State = adStateOpen Then rsCustomer.Close
    If rsJobDesc.State = adStateOpen Then rsJobDesc.Close
    If rsMechanic.State = adStateOpen Then rsMechanic.Close
End Sub

Private Sub mnuAdd_Click()
    If flxJob.TextMatrix(ActRow, 1) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    Else
        flxJob.AddItem ""
    End If
End Sub


Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxJob.Rows = 2 And ActRow = 1 Then
            flxJob.Rows = 1
            flxJob.Rows = 2
        Else
            flxJob.RemoveItem ActRow
        End If
        Call ComputeTotals
    End If
End Sub


Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    With flxMC
        For i = 0 To .Cols - 1
            .ColWidth(i) = .Width * 0.164
        Next i
        .Height = 315 * .Rows + 100
        lblHeading(4).Top = .Top + .Height + 50
        lblHeading(5).Top = lblHeading(4).Top + 30
    End With
    With flxJob
        .Col = 1: .Row = 0: .CellAlignment = 4
        .Top = lblHeading(4).Top + lblHeading(4).Height + 80
        .Height = Frame2.Top - (.Top + 100)
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 350) * 0.52
        .ColWidth(2) = (.Width - 350) * 0.12
        .ColWidth(3) = (.Width - 350) * 0.12
        .ColWidth(4) = (.Width - 350) * 0.12
        .ColWidth(5) = (.Width - 350) * 0.12
        txtTotal.Top = .Top + .Height
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4)
        txtTotal.Width = .ColWidth(5)
        lblTotal.Top = txtTotal.Top + 50
        lblTotal.Left = txtTotal.Left - (lblTotal.Width)
        txtDiscAmt.Top = txtTotal.Top + txtTotal.Height
        txtDiscAmt.Left = txtTotal.Left
        txtDiscAmt.Width = txtTotal.Width
        lblDisc.Top = txtDiscAmt.Top + 50
        lblDisc.Left = lblTotal.Left
        txtPctD.Top = txtDiscAmt.Top
        txtPctD.Left = lblDisc.Left - (txtPctD.Width + 50)
        txtGrandTtl.Top = txtDiscAmt.Top + txtDiscAmt.Height
        txtGrandTtl.Left = txtTotal.Left
        txtGrandTtl.Width = txtTotal.Width
        lblGrandTtl.Top = txtGrandTtl.Top + 50
        lblGrandTtl.Left = lblTotal.Left
        lblHeading(4).Left = .Left
        lblHeading(5).Left = .Left + 30
    End With
    lblHeading(0).Left = Frame1.Left
    lblHeading(1).Left = Frame1.Left + 30
    lblHeading(0).Top = Frame1.Top - (lblHeading(0).Height + 30)
    lblHeading(1).Top = Frame1.Top - lblHeading(1).Height
    lblHeading(2).Left = Frame1.Left + Frame1.Width - lblHeading(2).Width
    lblHeading(3).Left = lblHeading(2).Left + 30
    lblHeading(2).Top = Frame1.Top - (lblHeading(2).Height + 30)
    lblHeading(3).Top = Frame1.Top - lblHeading(3).Height
    
    
End Sub

Private Sub tmrDrop_Timer()
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
            cmdTrans(7).Left = cmdTrans(6).Width
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
        If BtnIndx = 5 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 5
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtAcct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr(KeyAscii)))
     If Chr(KeyAscii) = "'" Then KeyAscii = 0
     If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboMechanic_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxMC.SetFocus
        flxMC.Col = 0
    End If
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
                                           
Private Sub txtCustomer_Change()
    With flxCust
        If CustSrch Then
            CustSrch = False
            Exit Sub
        End If
        If txtCustomer.Text = "" Then .Visible = False: Exit Sub
        rsCustomer.Find "[CUSTOMER] LIKE '" + txtCustomer.Text + "%'", , adSearchForward, 1
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
            .Col = 0: .ColSel = 7
            CustSrch = True
            txtCustomer.Text = rsCustomer!CUSTOMER
            'Insert Additional retrieval here
            txtAddress.Text = flxCust.TextMatrix(flxCust.Row, 2)
            txtAcct.Text = flxCust.TextMatrix(flxCust.Row, 0)
            cboBrand.Text = flxCust.TextMatrix(flxCust.Row, 9) 'Brand
            txtCustID.Text = flxCust.TextMatrix(flxCust.Row, 11)    'CustomerID
            txtContact.Text = flxCust.TextMatrix(flxCust.Row, 3)
            flxMC.TextMatrix(1, 0) = flxCust.TextMatrix(flxCust.Row, 4) 'Model
            flxMC.TextMatrix(1, 1) = flxCust.TextMatrix(flxCust.Row, 5) 'Engine
            flxMC.TextMatrix(1, 2) = flxCust.TextMatrix(flxCust.Row, 6) 'chassis
            flxMC.TextMatrix(1, 3) = flxCust.TextMatrix(flxCust.Row, 7) 'Plate
            flxMC.TextMatrix(1, 4) = flxCust.TextMatrix(flxCust.Row, 8) 'type of Usage
            'Access
            'If flxCust.TextMatrix(flxCust.Row, 10) = "True" Then
            If flxCust.TextMatrix(flxCust.Row, 10) = "1" Then
            
                chkOnAcct.Value = 1 'On Acct
                
                'chkOnAcct.Visible = True
                txtDesignation.Text = flxCust.TextMatrix(flxCust.Row, 12) 'Designation
                chkType.Value = 1
            Else
                chkOnAcct.Value = 0 'On Acct
                chkType.Value = 0
                'chkOnAcct.Visible = False
                txtDesignation.Text = ""
            End If
            chkType.Visible = True
            cmdNewMC.Visible = True
            chkAddMC.Value = 0
            txtCustomer.SelStart = Val(txtCustomer.Tag)
            txtCustomer.SelLength = Len(txtCustomer.Text)
            CustBln = False
        Else
            'rsCustomer.Requery
            txtAddress.Text = ""
            txtAcct.Text = ""
            txtCustID.Text = ""
            txtContact.Text = ""
            txtAddress.Locked = False
            flxMC.TextMatrix(1, 0) = "" 'Model
            flxMC.TextMatrix(1, 1) = "" 'Engine
            flxMC.TextMatrix(1, 2) = "" 'chassis
            flxMC.TextMatrix(1, 3) = "" 'Plate
            chkOnAcct.Value = 0 'On Acct
            chkOnAcct.Visible = False
            chkType.Visible = False
            cmdNewMC.Visible = False
            chkAddMC.Value = 0
            chkType.Value = 0
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
            If rsCustomer.State = adStateOpen And rsCustomer.RecordCount > 1 Then rsCustomer.MoveFirst
        Case vbKeyBack
            CustSrch = True
            If rsCustomer.State = adStateOpen And rsCustomer.RecordCount > 1 Then rsCustomer.MoveFirst
    End Select

End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        flxCust.Visible = False
        txtAddress.SetFocus
    End If
End Sub

Private Sub txtCustomer_LostFocus()
    Dim rsAllMC As New ADODB.Recordset
    Dim Qry As String
    
    If Val(txtAcct.Text) = 0 Then Exit Sub
    '***Access
    'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CONTACTNO, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.TYPEOFUSAGE, SRVCUSTOMER.BRAND, SRVCUSTOMER.ONACCOUNT, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.ADDTLMC " & _
          "From SRVCUSTOMER WHERE (((SRVCUSTOMER.CUSTOMER)='" & txtCustomer.Text & "')) ORDER BY SRVCUSTOMER.CUSTOMER;"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CONTACTNO, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.TYPEOFUSAGE, SRVCUSTOMER.BRAND, SRVCUSTOMER.ONACCOUNT, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.ADDTLMC " & _
          "From SRVCUSTOMER WHERE SRVCUSTOMER.CUSTOMER='" & txtCustomer.Text & "' ORDER BY SRVCUSTOMER.CUSTOMER;"
    
    With rsAllMC
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxAllMC.Rows = 1
        cboEngine.Clear
        If .RecordCount > 1 Then
            .MoveFirst
            Do While Not .EOF
                flxAllMC.AddItem !ACCOUNTNO & vbTab & !CUSTOMER & vbTab & !Address & vbTab & IIf(IsNull(!CONTACTNO), "", !CONTACTNO) & vbTab & !MODEL & vbTab & !ENGINE & vbTab & !CHASSIS & vbTab & IIf(IsNull(!PLATENO), "", !PLATENO) & vbTab & IIf(IsNull(!TYPEOFUSAGE), "", !TYPEOFUSAGE) & vbTab & IIf(IsNull(!BRAND), "", !BRAND) & vbTab & IIf(IsNull(!ONACCOUNT), "", !ONACCOUNT) & vbTab & IIf(IsNull(!CUSTOMERID), "", !CUSTOMERID) & vbTab & IIf(IsNull(!DESIGNATION), "", !DESIGNATION)
                cboEngine.AddItem !ENGINE
                .MoveNext
            Loop
        End If
        .Close
        If cboEngine.ListCount > 1 Then
            lblAvlMC.Visible = True
            lblAvlMC.Caption = cboEngine.ListCount & " MC"
        Else
            lblAvlMC.Visible = False
        End If
    End With
End Sub

Private Sub txtDateProm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDateRec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDiscAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If ActCol <> 2 And KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        txtGrandTtl.SetFocus
    End If
End Sub

Private Sub txtDiscAmt_LostFocus()
    Dim Pctd As Single, myTtl As Single
    
    myTtl = Val(Replace(txtTotal.Text, ",", ""))
    If myTtl = 0 Then
        MsgBox "No Total Amount specified!", vbExclamation + vbOKOnly, "Missing Total"
        Exit Sub
    End If
    If Val(txtDiscAmt.Text) > myTtl Then
        MsgBox "Discount amount must not be greater than Total Amount!", vbExclamation + vbOKOnly, "Invalid Entry"
        txtDiscAmt.Text = ""
        Exit Sub
    End If
    If Val(txtDiscAmt.Text) > 0 Then
        Pctd = Val(txtDiscAmt.Text) / myTtl
        txtPctD.Text = Format(Pctd, "##0.00%")
        txtGrandTtl.Text = Format(myTtl - Val(txtDiscAmt.Text), "#,##0.00")
        If Val(txtDiscAmt.Text) < myTtl Then
            txtRemarks.Text = Replace(txtRemarks.Text, "FREE SERVICE", "")
        End If
        txtDiscAmt.Text = Format(txtDiscAmt.Text, "#,##0.00")
    ElseIf Val(txtDiscAmt.Text) = 0 Then
        txtPctD.Text = ""
        txtGrandTtl.Text = Format(myTtl, "#,##0.00")
    End If
End Sub

Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtDiscount.Text = Empty
        flxOrd.TextMatrix(ActRow, 2) = txtDiscount.Tag
        txtDiscount.Tag = "skiplostfocus"
        flxJob.SetFocus
        flxJob.Col = 4
        txtDiscount.Visible = False
    End If
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxJob.SetFocus
        If ActRow = flxJob.Rows - 1 Then flxJob.AddItem ""
        flxJob.Row = ActRow + 1
        flxJob.Col = 1
    End If
End Sub

Private Sub txtDiscount_LostFocus()
    If Val(txtDiscount.Text) > 100 Then
        MsgBox "Please enter a maximum discount of 100 only!", vbExclamation + vbOKOnly, "Validation Error"
        txtDiscount.Visible = False
        Exit Sub
    End If
    If Val(txtDiscount.Text) = 0 Then
        flxJob.TextMatrix(ActRow, 5) = flxJob.TextMatrix(ActRow, 3)
        flxJob.TextMatrix(ActRow, 4) = ""
    Else
        flxJob.TextMatrix(ActRow, 4) = Format(txtDiscount.Text, "##0.00") & "%"
        Call ComputeTotals
    End If
    txtDiscount.Visible = False
End Sub

Private Sub txtDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFlx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtDiscount.Text = Empty
        flxOrd.TextMatrix(ActRow, 2) = txtDiscount.Tag
        txtDiscount.Tag = "skiplostfocus"
        flxJob.SetFocus
        flxJob.Col = 4
        txtDiscount.Visible = False
    End If
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxJob.SetFocus
        flxJob.Col = 4
    End If
End Sub

Private Sub txtFlx_LostFocus()
    Dim MyTotal As Single, Ttl3 As Single, Ttl4 As Single, Ttl5 As Single
    
    flxJob.TextMatrix(ActRow, 3) = Format(txtFlx.Text, "#,##0.00")
    Call ComputeTotals
    txtFlx.Visible = False
End Sub

Public Sub ComputeTotals()
    Dim i As Integer
    Dim MyTotal As Single, DiscAmt As Single, Pctd As Single
    Dim JobTtl As Single, myDisc As Single, MyLabor As Single
        
    MyTotal = 0
    For i = 1 To flxJob.Rows - 1
        myDisc = Val(Replace(flxJob.TextMatrix(i, 4), "%", ""))
        MyLabor = Val(Replace(flxJob.TextMatrix(i, 3), ",", ""))
        JobTtl = MyLabor - (MyLabor * (myDisc / 100))
        flxJob.TextMatrix(i, 5) = Format(JobTtl, "#,##0.00")
        MyTotal = MyTotal + JobTtl
    Next i
    txtTotal.Text = Format$(MyTotal, "###,##0.00")
    If txtPctD.Text <> "" Then
        Pctd = Val(Replace(txtPctD.Text, "%", ""))
        DiscAmt = MyTotal * (Pctd / 100)
        txtDiscAmt.Text = Format(DiscAmt, "#,##0.00")
    ElseIf txtDiscAmt.Text <> "" Then
        DiscAmt = Val(Replace(txtDiscAmt.Text, ",", ""))
    Else
        DiscAmt = 0
    End If
    txtGrandTtl.Text = Format$(MyTotal - DiscAmt, "#,##0.00")
    txtBal.Text = txtGrandTtl.Text
End Sub

Private Sub txtFSC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtFSC.Text = Empty
        flxOrd.TextMatrix(ActRow, 2) = txtFSC.Tag
        txtFSC.Tag = "skiplostfocus"
        flxJob.SetFocus
        flxJob.Col = 2
        txtFSC.Visible = False
    End If
End Sub
Private Sub txtFSC_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxJob.SetFocus
        flxJob.Col = 3
    End If
End Sub

Private Sub txtFSC_LostFocus()
    flxJob.TextMatrix(ActRow, 2) = txtFSC.Text
    If txtFSC.Text <> "" Then
        flxJob.TextMatrix(ActRow, 3) = ""
        flxJob.TextMatrix(ActRow, 4) = ""
        flxJob.TextMatrix(ActRow, 5) = ""
        Call ComputeTotals
    End If
    txtFSC.Visible = False
End Sub

Private Sub txtJobNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtJobNo_LostFocus()
    Dim xQry As String
    
    If Not AddNew Then Exit Sub
    '***Access
    'xQry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER " & _
           "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtJobNo.Text & "'));"
    'mySQl
    xQry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER " & _
           "WHERE SRVCJOBORDER.JOBORDERNO='" & Trim(txtJobNo.Text) & "';"
    
    If CtrlNumExist(xQry) Then
        MsgBox "Job Order No. " & txtJobNo.Text & " already exist!", vbExclamation + vbOKOnly, "Exist"
        txtJobNo.SetFocus
        txtJobNo.SelStart = 0
        txtJobNo.SelLength = Len(txtJobNo.Text)
    End If
End Sub

Private Sub cboUsage_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call cboUsage_LostFocus
        flxMC.SetFocus
        flxMC.Col = 5
    End If
End Sub

Private Sub cboUsage_LostFocus()
    If cboUsage.Tag = "skiplostfocus" Then
        cboUsage.Tag = ""
        Exit Sub
    End If
    flxMC.TextMatrix(MCrow, MCcol) = cboUsage.Text
    cboUsage.Visible = False
End Sub

Private Sub cboUsage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxMC.TextMatrix(MCrow, MCcol) = cboUsage.Tag
        cboUsage.Tag = "skiplostfocus"
        cboUsage.Visible = False
        flxMC.SetFocus
        flxMC.Col = 5
    End If
End Sub

Private Sub txtMC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxMC.TextMatrix(ActRow, ActCol) = txtMC.Tag
        txtMC.Tag = "skiplostfocus"
        txtMC.Visible = False
        flxMC.SetFocus
        flxMC.Col = MCcol
    End If
End Sub

Private Sub txtMC_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call txtMC_LostFocus
        flxMC.SetFocus
        If MCcol = 5 Then
            flxJob.SetFocus
            flxJob.Col = 1
            flxJob.Row = 1
        Else
            flxMC.Col = MCcol + 1
        End If
    End If
End Sub

Private Sub txtMC_LostFocus()
    Dim xQry As String, vMsg As String
    
    If txtMC.Tag = "skiplostfocus" Then
        txtMC.Tag = ""
        Exit Sub
    End If
    If MCcol = 1 Then 'search existing engine
        If txtAcct.Text = "" Then
            'xQry = "SELECT SRVCUSTOMER.ENGINE FROM SRVCUSTOMER WHERE (((SRVCUSTOMER.ENGINE)='" & txtMC.Text & "'));"
            'mySQL
            xQry = "SELECT SRVCUSTOMER.ENGINE FROM SRVCUSTOMER WHERE SRVCUSTOMER.ENGINE='" & txtMC.Text & "';"
            If CtrlNumExist(xQry) Then
                vMsg = GetEngineOwner(txtMC.Text)
                MsgBox "Engine No. " & flxMC.TextMatrix(1, MCcol) & " already exist!" & vbCrLf & vMsg, vbExclamation + vbOKOnly, "Exist"
                flxMC.TextMatrix(MCrow, MCcol) = ""
                txtMC.Visible = False
            End If
        End If
    End If
    flxMC.TextMatrix(MCrow, MCcol) = txtMC.Text
    txtMC.Visible = False
End Sub

Function GetEngineOwner(MyEN As String) As String
    Dim rsEngine As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE (((SRVCUSTOMER.ENGINE)='" & MyEN & "'));"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE SRVCUSTOMER.ENGINE='" & MyEN & "';"
    
    With rsEngine
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetEngineOwner = "Customer: " & !CUSTOMER & vbCrLf & "Address: " & !Address & vbCrLf & "Account No. " & !ACCOUNTNO
        Else
            GetEngineOwner = ""
        End If
        .Close
    End With
    
End Function

Private Sub txtPctD_GotFocus()
    If txtPctD.Text <> "" Then
        txtPctD.SelStart = 0
        txtPctD.SelLength = Len(txtPctD)
    End If
End Sub

Private Sub txtPctD_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If ActCol <> 2 And KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        txtGrandTtl.SetFocus
    End If
End Sub

Private Sub txtPctD_LostFocus()
    Dim DiscAmt As Single, myTtl As Single
    Dim OldDisc As String
    
    OldDisc = txtPctD.Text
    If InStr(txtPctD.Text, "%") > 0 Then txtPctD.Text = Replace(txtPctD.Text, "%", "")
    
    myTtl = Val(Replace(txtTotal.Text, ",", ""))
    If myTtl = 0 Then
        MsgBox "No Total Amount specified!", vbExclamation + vbOKOnly, "Missing Total"
        Exit Sub
    End If
    If Val(txtPctD.Text) > 100 Then
        MsgBox "Percent Discount must not exceed 100!", vbExclamation + vbOKOnly, "Invalid Entry"
        txtPctD.Text = ""
        Exit Sub
    End If
    If Val(txtPctD.Text) > 0 Then
        DiscAmt = myTtl * (Val(txtPctD.Text) / 100)
        If Val(txtPctD.Text) = 100 Then
            If InStr(txtRemarks.Text, "FREE SERVICE") = 0 Then
                txtRemarks.Text = "FREE SERVICE " & txtRemarks.Text
            End If
        Else
            txtRemarks.Text = Replace(txtRemarks.Text, "FREE SERVICE", "")
        End If
        txtDiscAmt.Text = Format(DiscAmt, "#,##0.00")
        txtPctD.Text = Format(txtPctD.Text, "##0.00") & "%"
        txtGrandTtl.Text = Format(myTtl - DiscAmt, "#,##0.00")
    Else
        txtDiscAmt.Text = "0.00"
        txtGrandTtl.Text = Format(myTtl, "#,##0.00")
        txtRemarks.Text = Replace(txtRemarks.Text, "FREE SERVICE", "")
    End If
    
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

