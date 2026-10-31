VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form arPR 
   Caption         =   "Provisional Receipt"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   10995
   WindowState     =   2  'Maximized
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
      Left            =   5160
      TabIndex        =   39
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   6120
      TabIndex        =   38
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   6120
      TabIndex        =   37
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboAccount 
      Height          =   315
      Left            =   840
      TabIndex        =   36
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   120
      TabIndex        =   35
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid flxAOP 
      Height          =   2775
      Left            =   120
      TabIndex        =   11
      Tag             =   "XY"
      Top             =   3840
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   $"ProvReceiptold.frx":0000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1695
      Left            =   2520
      TabIndex        =   27
      Top             =   7800
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
      Width           =   10695
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
         Left            =   8040
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
         Left            =   8040
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   575
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
         Left            =   8040
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
         Left            =   8040
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
         Left            =   6720
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   1372
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
         Left            =   6720
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   999
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
         Left            =   6720
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   627
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
         Left            =   6720
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
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   720
      Width           =   10695
      Begin MSMask.MaskEdBox txtPRDate 
         Height          =   315
         Left            =   8880
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   585
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
         Left            =   8880
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
         Left            =   7920
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
         Left            =   7680
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
      Image           =   "ProvReceiptold.frx":009B
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
      Image           =   "ProvReceiptold.frx":0635
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
      Image           =   "ProvReceiptold.frx":0BCF
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
      Image           =   "ProvReceiptold.frx":1169
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
      Image           =   "ProvReceiptold.frx":1703
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
      Image           =   "ProvReceiptold.frx":1C9D
      cBack           =   -2147483633
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
      Top             =   3600
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
      Left            =   7320
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
      Left            =   7350
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
      Left            =   5400
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11115
      m_MDI_Height    =   8220
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
Dim CustSrch As Boolean, CustBln As Boolean, AddNew As Boolean
Dim ActRow As Integer, ActCol As Integer

Private Sub ClearEntries()
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
End Sub

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    Frame2.Enabled = Not Bln
    flxAOP.Enabled = Not Bln
End Sub

Private Sub tbSaveCancel(Bln As Boolean)
    Dim i As Integer
    
    If Not Bln Then
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
            flxAOP.TextMatrix(ActRow, 3) = rsAcctCust!Address
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
    cboAccount.Visible = False
    txtRemarks.Visible = False
    txtAmount.Visible = False
End Sub


Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Function GetPRNo() As String
    Dim rsPRno As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    Qry = "SELECT Max(CHECKS.PRNO) AS MaxOfPRNO FROM CHECKS;"

    With rsPRno
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, .Fields(0) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    GetPRNo = Format(MyNo, "0000")
End Function


Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer
    Dim eMsg As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=46));"
            'mySQL Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=46;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            Call ClearEntries
            Call LockData(False)
            Call tbSaveCancel(True)
            txtCustomer.SetFocus
            txtPRDate.Text = Format$(Date, "mm/dd/yy")
            txtPRno.Text = GetPRNo
            AddNew = True
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=46));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=46;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtChksPRid.Text) = 0 Then
                MsgBox "No available P.R. to edit!", vbExclamation + vbOKOnly, "Edit"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Edit PR?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call LockData(False)
            Call tbSaveCancel(True)
            txtCustomer.SetFocus
            txtChksPRid.Tag = txtPRno.Text
            AddNew = False
        Case 2  'Summary
            arCheckSumm.Show
            Unload Me
        Case 3  'Quit
            Unload Me
            MotoRep.tbMain.Visible = True
        Case 4  'Save
            eMsg = ValidatePR
            If eMsg <> "" Then
                MsgBox eMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save PR?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'Call Save Check
            Call SaveCheck
            MsgBox "P.R. successfully saved!", vbOKOnly + vbInformation, "Save"
            Call LockData(True)
            Call tbSaveCancel(False)
            
        Case 5  'Cancel
            Call HideObj
            Ans = MsgBox("Do you want to Cancel PR?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            Call ClearEntries
            Call LockData(True)
            Call tbSaveCancel(False)
    End Select
        
End Sub

Public Sub ShowRecord(MyId As Long)
    Dim rsCheck As New ADODB.Recordset
    Dim rsAlloc As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT CHECKS.* From CHECKS WHERE (((CHECKS.CHKSPRID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT CHECKS.* From CHECKS WHERE CHECKS.CHKSPRID=" & MyId & ";"

    With rsCheck
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            CustSrch = True
            txtCustomer.Text = IIf(IsNull(!CUSTOMER), "", !CUSTOMER)
            txtAddress.Text = IIf(IsNull(!Address), "", !Address)
            txtPRno.Text = IIf(IsNull(!PRno), "", !PRno)
            If Not IsNull(!PRDATE) Then txtPRDate.Text = Format(!PRDATE, "mm/dd/yy")
            txtCustID.Text = !CUSTOMERID
            txtAcctName.Text = !ChkAcctName
            txtBank.Text = !Bank
            txtChkNo.Text = !CheckNo
            txtChkAmt.Text = Format(!Amount, "#,##0.00")
            If !Status = "FOR DEPOSIT" Then
                cboChkStat.Text = "CURRENT-FOR DEPOSIT"
            Else
                cboChkStat.Text = IIf(IsNull(!Status), "", !Status)
            End If
            txtPaidTo.Text = !PAIDTO
            txtChkDate.Text = Format(!CheckDate, "mm/dd/yy")
            cboChkType.Text = IIf(IsNull(!CheckType), "", !CheckType)
            txtChksPRid.Text = !CHKSPRID
        End If
        .Close
    End With
    'Insert Allocated Payments
    
    '***Access
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CONTACTPERSON, CUSTOMER.ADDRESS, CUSTOMER.OSBAL, CHECKALLOC.CUSTOMERID, CHECKALLOC.CHKSPRID " & _
          "FROM CHECKALLOC INNER JOIN CUSTOMER ON CHECKALLOC.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((CHECKALLOC.CHKSPRID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CONTACTPERSON, CUSTOMER.ADDRESS, CUSTOMER.OSBAL, CHECKALLOC.CUSTOMERID, CHECKALLOC.CHKSPRID " & _
          "FROM CHECKALLOC INNER JOIN CUSTOMER ON CHECKALLOC.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE CHECKALLOC.CHKSPRID=" & MyId & ";"
    
    With rsAlloc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxAOP.Rows = 1
            Do While Not .EOF
                flxAOP.AddItem vbTab & !CustomerName & vbTab & !ContactPerson & vbTab & !Address & vbTab & Format(!OSbal, "#,##0.00") & vbTab & !CUSTOMERID
                .MoveNext
            Loop
        End If
        .Close
    End With
End Sub
Private Sub SaveCheck()
    Dim rsCheck As New ADODB.Recordset
    Dim rsAlloc As New ADODB.Recordset
    Dim ChkAmt As Currency, i As Integer, Qry As String
    
    'With rsCheck
    '    If AddNew Then
    '        .Open "Checks", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
    '        Qry = "SELECT CHECKS.* From CHECKS WHERE (((CHECKS.CHKSPRID)=" & txtChksPRid.Text & "));"
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '
    '        If .RecordCount > 0 Then .MoveFirst
    '    End If
    '    !CUSTOMER = txtCustomer.Text
    '    !ADDRESS = txtAddress.Text
    '    !PRNO = txtPRno.Text
    '    !PRDATE = txtPRDate.Text
    '    !CUSTOMERID = Val(txtCustID.Text)
    '    !CHKACCTNAME = txtAcctName.Text
    '    !BANK = txtBank.Text
    '    !CHECKNO = txtChkNo.Text
    '    ChkAmt = Val(Replace(txtChkAmt.Text, ",", ""))
    '    !AMOUNT = ChkAmt
    '    !Status = cboChkStat.Text
    '    If cboChkStat.Text = "CURRENT-FOR DEPOSIT" Then
    '        !DATE4DPSIT = txtPRDate.Text
    '        !Status = "FOR DEPOSIT"
    '    ElseIf cboChkStat.Text = "CURRENT-CLEARED" Then
    '        !DATE4BRKDWN = txtPRDate.Text
    '    End If
    '    !PAIDTO = txtPaidTo.Text
    '    !CHECKDATE = txtChkDate.Text
    '    !CHECKTYPE = cboChkType.Text
    '    If AddNew Then !CHECKBAL = ChkAmt
    '    !UPDATEDBY = UserID
    '    !LASTUPDATE = Now
    '    .Update
    '    txtChksPRid.Text = !CHKSPRID
    '    .Close
    'End With
    
    'mySQL
    If cboChkStat.Text = "CURRENT-FOR DEPOSIT" Then
        myDATE4DPSIT = txtPRDate.Text
        myStatus = "FOR DEPOSIT"
    ElseIf cboChkStat.Text = "CURRENT-CLEARED" Then
        myDATE4BRKDWN = txtPRDate.Text
    End If
    ChkAmt = Val(Replace(txtChkAmt.Text, ",", ""))
    
    If AddNew Then
        LYRcon.Execute "Insert Into checks(CUSTOMER, ADDRESS, PRNO, PRDATE, CUSTOMERID, CHKACCTNAME, BANK, " & _
            "CHECKNO, AMOUNT, `STATUS`, PAIDTO, CHECKDATE, CHECKTYPE, CHECKBAL, UPDATEDBY, LASTUPDATE)" & _
            "VALUES('" & txtCustomer.Text & "', '" & txtAddress.Text & "', '" & txtPRno.Text & "', " & _
            "'" & Format(txtPRDate.Text, "yyyy-mm-dd") & "', " & Val(txtCustID.Text) & ", '" & txtAcctName.Text & "', " & _
            "'" & txtBank.Text & "', '" & txtChkNo.Text & "', " & ChkAmt & ", '" & cboChkStat.Text & "', '" & txtPaidTo.Text & "', " & _
            "'" & Format(txtChkDate.Text, "yyyy-mm-dd") & "', '" & cboChkType.Text & "', " & ChkAmt & ", " & UserID & ", " & _
            "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
    Else
        LYRcon.Execute "Update checks SET CUSTOMER = '" & txtCustomer.Text & "', ADDRESS = '" & txtAddress.Text & "', " & _
            "PRNO = '" & txtPRno.Text & "', PRDATE = '" & Format(txtPRDate.Text, "yyyy-mm-dd") & "', CUSTOMERID = " & Val(txtCustID.Text) & ", " & _
            "CHKACCTNAME = '" & txtAcctName.Text & "', BANK = '" & txtBank.Text & "', CHECKNO = '" & txtChkNo.Text & "', " & _
            "AMOUNT = " & ChkAmt & ", `STATUS` = '" & cboChkStat.Text & "', PAIDTO = '" & txtPaidTo.Text & "', " & _
            "CHECKDATE = '" & Format(txtChkDate.Text, "yyyy-mm-dd") & "', CHECKTYPE = '" & cboChkType.Text & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
            "Where checks.CHKSPRID = " & txtChksPRid.Text & ";"
   End If
    
    'Update Check Allocation
   If cboChkType.Text = "CUSTOMER CHECK" Then
       If Not AddNew Then
           '***Access
           'LYRcon.Execute "DELETE CHECKALLOC.*, CHECKALLOC.CHKSPRID From CHECKALLOC WHERE (((CHECKALLOC.CHKSPRID)=" & txtChksPRid.Text & "));"
           'mySQL
           LYRcon.Execute "DELETE From CHECKALLOC WHERE CHECKALLOC.CHKSPRID=" & txtChksPRid.Text & ";"
       End If
       'With rsAlloc
       '    .Open "CheckAlloc", LYRcon, adOpenKeyset, adLockOptimistic
       '    For i = 1 To flxAOP.Rows - 1
       '        .AddNew
       '        !CHKSPRID = txtChksPRid.Text
       '        !CUSTOMERID = Val(flxAOP.TextMatrix(i, 5))
       '        .Update
       '    Next i
       '    .Close
       'End With
       For i = 1 To flxAOP.Rows - 1
           LYRcon.Execute "Insert Into checkalloc( CHKSPRID, CUSTOMERID)" & _
               "VALUES( " & txtChksPRid.Text & ", " & Val(flxAOP.TextMatrix(i, 5)) & ");"
       Next i
   End If
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
        If .TextMatrix(.Rows - 1, 5) = "" And .Rows > 2 Then .RemoveItem .Rows - 1
        If .TextMatrix(.Rows - 1, 5) = "" And cboChkType.Text = "CUSTOMER CHECK" Then vMsg = vMsg & vbCrLf & "No Customer allocated for this Check"
        If cboChkType.Text <> "RE-DISCOUNT" Then
            If .TextMatrix(.Rows - 1, 5) = "" And Val(txtCustID.Text) = 0 Then vMsg = vMsg & vbCrLf & "You are required to allocate this Check to a Customer"
        End If
    End With

    ValidatePR = vMsg

End Function

Private Sub flxAOP_Click()
    ActRow = flxAOP.Row
    ActCol = flxAOP.Col
    
    If ActCol = 0 Or ActRow = 0 Or Val(txtCustID.Text) > 0 Then Exit Sub
    Select Case ActCol
        Case 1    'Account No
            With cboAccount
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
    End Select
End Sub

Private Sub flxAOP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxAOP_Click
End Sub

Private Sub flxAOP_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
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
    Timer1.Enabled = True
    Call tbSaveCancel(False)
    Call LockData(True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsCustomer.State = adStateOpen Then rsCustomer.Close
    If rsAcctCust.State = adStateOpen Then rsAcctCust.Close
End Sub

Private Sub mnuAdd_Click()
    'validate Last Row
    With flxAOP
        If .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 4) = "" Then
            MsgBox "Incomplete entry of Last Row!", vbExclamation + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        .AddItem ""
    End With
End Sub

Private Sub mnuDelete_Click()
    If ActRow = 1 And flxAOP.Rows = 2 Then
        flxAOP.Rows = 1
        flxAOP.Rows = 2
    Else
        flxAOP.RemoveItem ActRow
    End If
    Call ComputeTotal
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
        .ColWidth(1) = (.Width - 300) * 0.24
        .ColWidth(2) = (.Width - 300) * 0.24
        .ColWidth(3) = (.Width - 300) * 0.3
        .ColWidth(4) = (.Width - 300) * 0.2
        .ColWidth(5) = 0
        
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
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON, CUSTOMER.OSBAL, CUSTOMER.CTYPE " & _
          "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
    With rsAcctCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboAccount.Clear
            Do While Not .EOF
                cboAccount.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
    With cboChkType
        .AddItem "CUSTOMER CHECK"
        .AddItem "PETTY CASH"
        .AddItem "RE-DISCOUNT"
        .AddItem "OTHER CHECKS"
    End With
    With cboChkStat
        .AddItem "CURRENT-CLEARED"
        .AddItem "CURRENT-FOR DEPOSIT"
        .AddItem "POST-DATED CHECK"
    End With
    Timer1.Enabled = False
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
    
    strValid = "123456789.0"
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
    Dim myTtl As Single, myChkAmt As Single
    Dim i As Integer
    
    myTtl = 0
    For i = 1 To flxAOP.Rows - 1
        myTtl = myTtl + PureNum(flxAOP.TextMatrix(i, 4))
    Next i
    myChkAmt = PureNum(txtChkAmt.Text)
    txtTtlPay.Text = FormatNumber(myTtl, 2, vbTrue)
    If myChkAmt > myTtl Then
        txtRefund.Text = FormatNumber(myChkAmt - myTtl, 2, vbTrue)
        txtRefund.Visible = True
        lblRefund.Visible = True
    Else
        txtRefund.Text = ""
        txtRefund.Visible = False
        lblRefund.Visible = False
    End If
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
    txtPaidTo.Text = "MOTO-R REPUBLIC CORP."
    cboChkType.ListIndex = 0
    If Val(txtCustID.Text) > 0 Then
        ActRow = 1
        cboAccount.Text = txtCustomer.Text
        flxAOP.TextMatrix(1, 1) = txtCustomer.Text
    End If
End Sub

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
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtPRDate.SetFocus
End Sub
Private Sub txtPRno_LostFocus()
    If txtPRno.Text <> "" Then
        '***Access
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
