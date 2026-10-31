VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form CashFlowMgrValidation 
   Caption         =   "Validation Slip Confirmation"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTabNewBank 
      Height          =   2775
      Left            =   10200
      TabIndex        =   59
      Top             =   2760
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "NEW BANK ACCOUNT"
      TabPicture(0)   =   "CashFlowMgrValidation.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOK"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame8 
         Height          =   1695
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   6255
         Begin VB.TextBox txtNewBankName 
            Height          =   375
            Left            =   1680
            TabIndex        =   62
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox txtNewBankAcctNo 
            Height          =   375
            Left            =   1680
            TabIndex        =   63
            Top             =   720
            Width           =   4455
         End
         Begin VB.TextBox txtNewBankAcctName 
            Height          =   375
            Left            =   1680
            TabIndex        =   64
            Top             =   1200
            Width           =   4455
         End
         Begin VB.Label Label19 
            Caption         =   "BANK NAME:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "BANK ACCT NAME:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "BANK ACCT #:"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   1095
         End
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   495
         Left            =   2040
         TabIndex        =   65
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         Image           =   "CashFlowMgrValidation.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3360
         TabIndex        =   66
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         Image           =   "CashFlowMgrValidation.frx":05B6
         cBack           =   -2147483633
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   3
      Left            =   5340
      TabIndex        =   25
      Top             =   0
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   741
      Caption         =   "Exit Report"
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
      Image           =   "CashFlowMgrValidation.frx":0B50
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   2
      Left            =   3600
      TabIndex        =   27
      Top             =   0
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   741
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
      Image           =   "CashFlowMgrValidation.frx":10EA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   1
      Left            =   1860
      TabIndex        =   26
      Top             =   0
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   741
      Caption         =   "BOF/Cash-Out"
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
      cFHover         =   255
      LockHover       =   2
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "CashFlowMgrValidation.frx":1684
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   741
      Caption         =   "&Add Deposit Slip"
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
      Image           =   "CashFlowMgrValidation.frx":1C1E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   4
      Left            =   10200
      TabIndex        =   28
      Top             =   360
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   741
      Caption         =   "Save"
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
      Image           =   "CashFlowMgrValidation.frx":21B8
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   10200
      TabIndex        =   20
      Text            =   "HoldBackColor"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Managers Cash FLow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   9855
      Begin MSComCtl2.DTPicker DTMgrDate 
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   103612417
         CurrentDate     =   43327
      End
      Begin VB.TextBox txtcheque 
         Height          =   375
         Left            =   7200
         TabIndex        =   16
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtcash 
         Height          =   375
         Left            =   7200
         TabIndex        =   14
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtprpby 
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label14 
         Caption         =   "Check Deposit:"
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Cash Deposit:"
         Height          =   255
         Left            =   6000
         TabIndex        =   13
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Mgr.Cash Flow Prepared by:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Mgr. Cash Flow Date:"
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   510
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9855
      Begin VB.TextBox txtBankName 
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   200
         Width           =   3135
      End
      Begin VB.TextBox txtBankAcctno 
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   600
         Width           =   3135
      End
      Begin lvButton.lvButtons_H cmdNewBank 
         Height          =   375
         Left            =   4870
         TabIndex        =   58
         Top             =   240
         Width           =   480
         _ExtentX        =   847
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
         ImgAlign        =   1
         Image           =   "CashFlowMgrValidation.frx":2752
         ImgSize         =   40
         cBack           =   -2147483633
      End
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   5400
         TabIndex        =   22
         Top             =   120
         Width           =   4335
         Begin VB.Label lblValid 
            Alignment       =   2  'Center
            Caption         =   "UNVALIDATED"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.ComboBox cboBANKACCTNO 
         Height          =   315
         Left            =   5520
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox cboBANKNAME 
         Height          =   315
         Left            =   5520
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txtBANKACCTNAME 
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1000
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Bank Account Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Account No:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   680
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Bank Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1815
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   420
      Index           =   5
      Left            =   10200
      TabIndex        =   29
      Top             =   0
      Width           =   1740
      _ExtentX        =   3069
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
      Image           =   "CashFlowMgrValidation.frx":2CEC
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtChckNo 
      Height          =   375
      Left            =   10200
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCVNO 
      Height          =   375
      Left            =   10200
      TabIndex        =   32
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Validation Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   9855
      Begin VB.ComboBox cboValidType 
         Height          =   315
         Left            =   6720
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CheckBox chkCheque 
         Caption         =   "Apply to Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid flxCV 
         Height          =   2160
         Left            =   9480
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   3360
         Visible         =   0   'False
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3810
         _Version        =   393216
         Cols            =   6
         RowHeightMin    =   315
         BackColorBkg    =   14737632
         Enabled         =   -1  'True
         FormatString    =   "|^PCV NO |<CV DATE |<PARTICULARS |>AMOUNT"
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   9615
         Begin VB.OptionButton opt2 
            Caption         =   "Deposited in Bank by:"
            Height          =   255
            Left            =   4680
            TabIndex        =   38
            Top             =   285
            Width           =   1935
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Picked Up"
            Height          =   255
            Left            =   3600
            TabIndex        =   39
            Top             =   285
            Width           =   1815
         End
         Begin VB.TextBox txtactcash 
            Height          =   375
            Left            =   1560
            TabIndex        =   37
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtvalnum 
            Height          =   375
            Left            =   6600
            TabIndex        =   36
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtdepby 
            Height          =   375
            Left            =   6600
            TabIndex        =   35
            Top             =   240
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker txtDepositDate 
            Height          =   375
            Left            =   1560
            TabIndex        =   34
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   103153665
            CurrentDate     =   43328
         End
         Begin VB.Label Label8 
            Caption         =   "Date Deposit :"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   285
            Width           =   2175
         End
         Begin VB.Label lblActCash 
            Caption         =   "Deposited Amount:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Validation Number:"
            Height          =   255
            Left            =   5040
            TabIndex        =   40
            Top             =   780
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   9615
         Begin VB.TextBox txtBOFAmnt 
            Height          =   375
            Left            =   1560
            TabIndex        =   45
            Top             =   720
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker dtBOFDate 
            Height          =   375
            Left            =   1560
            TabIndex        =   46
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   103153665
            CurrentDate     =   43328
         End
         Begin VB.Label Label12 
            Caption         =   "Total BOF Amount:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Validation Date"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   285
            Width           =   1455
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxChecks 
         Height          =   2160
         Left            =   120
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   1800
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3810
         _Version        =   393216
         Cols            =   7
         RowHeightMin    =   315
         BackColorBkg    =   14737632
         Enabled         =   -1  'True
         FormatString    =   "|^CHECK NO. |<CHECK NAME |^ CHECK TYPE |> AMOUNT |<CHECK DATE"
      End
      Begin VB.Frame Frame7 
         Height          =   2175
         Left            =   120
         TabIndex        =   49
         Top             =   1800
         Width           =   9615
         Begin VB.TextBox txtPurpose 
            Height          =   405
            Left            =   960
            TabIndex        =   56
            Top             =   1440
            Width           =   8535
         End
         Begin VB.TextBox txtReason 
            Height          =   405
            Left            =   960
            TabIndex        =   55
            Top             =   840
            Width           =   8535
         End
         Begin VB.TextBox txtCVDATE 
            Height          =   375
            Left            =   6600
            TabIndex        =   57
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtCVNumber 
            Height          =   375
            Left            =   960
            TabIndex        =   54
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label18 
            Caption         =   "PURPOSE:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1515
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "REASON:"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "CV DATE:"
            Height          =   255
            Left            =   5760
            TabIndex        =   51
            Top             =   260
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "CV NO:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   260
            Width           =   615
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Validation Type:"
         Height          =   255
         Left            =   5160
         TabIndex        =   31
         Top             =   1440
         Width           =   1575
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   10320
      m_MDI_Height    =   8760
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "CashFlowMgrValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mainRs As New ADODB.Recordset
Dim strPartial As String, strTotal As String
Dim formAction As String, myValidNo As String
Dim Ans As String, Tmp() As String
Dim Qry As String
Dim myMsg As String
Dim i As Long, j As Long, myBANKID As Long, myMGRCFID As Long, myDEPSLIPID As Integer
Dim intTotalAmount As Currency
Dim arrBANK() As String
Dim arrBANKACCT() As String
Dim StatSrch As Boolean, AddNew As Boolean, ProperExit As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim myValidDate As Date

Public Sub getBankDetails(MyID As Long)
Dim rsBANK As New ADODB.Recordset

    If MyID = 0 Then
        Qry = "Call sp_mrcretails.sp_MgrCshFlow_getBankAcct('" & sysBranchCode & "')"
        With rsBANK
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                txtBankName.Text = IIf(IsNull(!BANKNAME), "", !BANKNAME)
                txtBankAcctno.Text = IIf(IsNull(!BANKACCTNO), "", !BANKACCTNO)
                txtBANKACCTNAME.Text = IIf(IsNull(!BANKACCTNAME), "", !BANKACCTNAME)
                myBANKID = IIf(IsNull(!BANKID), 0, !BANKID)
            End If
            .Close
        End With
    End If
End Sub
Sub clearform()
    'Frame 1
'    cboBANKNAME.Text = ""
'    cboBANKACCTNO.Text = ""
    txtBankName.Text = ""
    txtBankAcctno.Text = ""
    txtBANKACCTNAME.Text = ""
    
    'Frame 3
    txtdepby.Text = ""
    txtactcash.Text = 0#
    txtvalnum.Text = ""
    cboValidType.Text = ""
    chkCheque.Value = 0
    
    myBANKID = 0
    myDEPSLIPID = 0

    flxChecks.Rows = 1
    flxChecks.Rows = 2
    
    txtBOFAmnt.Text = 0#
    txtCVDATE.Text = ""
    txtCVNumber.Text = ""
    txtReason.Text = ""
    txtPurpose.Text = ""
    
    lblValid.Caption = "UNVALIDATED"
End Sub

Public Sub flxChanger()
    With flxChecks
        .Cols = 7
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.2   'CHECK NO
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.3  'CHECK NAME
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.2 'CHECK TYPE
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.14 'AMOUNT
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.12 'CHECK DATE
        .ColWidth(6) = 0 '(.Width - .ColWidth(0)) * 0.07 'CHKSPRID
    End With

End Sub

Sub CreateTempTable()
On Error GoTo er

        LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcretails.TMP_CHECKS;"

         DoEvents
        LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcretails.TMP_CHECKS( " & _
        "   DEPSLIPID       Int(11),        " & _
        "   CHECKNO         Varchar(500),   " & _
        "   CHECKTYPE       Varchar(500),   " & _
        "   AMOUNT          Decimal(10,2),  " & _
        "   VALIDNO         Varchar(500),   " & _
        "   CHECKNAME       Varchar(500),   " & _
        "   CHECKDATE       Date            " & _
        ");"
        
        Qry = "Insert into sp_mrcretails.TMP_CHECKS(CHECKNO, CHECKTYPE, AMOUNT, VALIDNO, CHECKNAME, CHECKDATE) VALUES "
        For i = 1 To flxChecks.Rows - 1
            If i > 1 Then Qry = Qry & ","
            Qry = Qry & "('" & flxChecks.TextMatrix(i, 1) & "','" & flxChecks.TextMatrix(i, 3) & "','" & Replace(flxChecks.TextMatrix(i, 4), ",", "") & "','" & Trim(txtvalnum.Text) & "','" & flxChecks.TextMatrix(i, 2) & "','" & Format(flxChecks.TextMatrix(i, 5), "YYYY-MM-DD") & "') "
        Next i
        
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Function ValidateDepositSlip() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
If cboValidType.Text <> "BOF/CASHOUT" Then
    vMsg = ""
    If cboValidType.Text = "CHECKS" Then
        If flxChecks.TextMatrix(flxChecks.Rows - 1, 1) = "" Then
            If flxChecks.Rows = 2 And cboValidType.Text = "CHECKS" Then
                flxChecks.Rows = 1
                flxChecks.Rows = 2
                vMsg = "No Checque Found!"
            ElseIf flxChecks.Rows = 2 Then
                flxChecks.Rows = 1
                flxChecks.Rows = 2
            Else
                flxChecks.RemoveItem flxChecks.Rows - 1
            End If
        End If
        
        If flxChecks.Rows > 2 Then
            If flxChecks.TextMatrix(flxChecks.Rows - 1, 1) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on Last Row"
        End If
    End If
    
    
    If Not IsDate(txtDepositDate.Value) Then vMsg = "Invalid Deposit Date"
    If myBANKID <= 0 Then
        vMsg = vMsg & vbCrLf & "No BANK or ACCOUNT NUMBER Selected"
    End If
    If opt2.Value = False And txtdepby.Text <> "" Then
        opt2.Value = True
        'vMsg = vMsg & vbCrLf & "No Deposited By Found!"
    End If
    If txtactcash.Text = "" Or Val(txtactcash.Text) <= 0 Then
        If cboValidType.Text = "CASH" Then
            vMsg = vMsg & vbCrLf & "No Cash Deposited Amount Encoded!"
        Else
            If flxChecks.Rows <= 2 Then
                vMsg = vMsg & vbCrLf & "No Cheque Found!"
            Else
                vMsg = vMsg & vbCrLf & "No Cash Deposited Amount Encoded!"
            End If
        End If
    End If
    If txtvalnum.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Validation Number"
    If txtBANKACCTNAME.Text = "" Then vMsg = vMsg & vbCrLf & "No Bank Account Name Found!"
    If cboValidType.Text = "" Then vMsg = vMsg & vbCrLf & "No Deposit Slip Type Found!"
    If cboValidType.ListIndex = -1 Then vMsg = vMsg & vbCrLf & "Invalid Deposit Slip Type! Please select from the list."
Else
    If txtReason.Text = "" Then vMsg = vMsg & vbCrLf & "No Reason Found!"
    If txtPurpose.Text = "" Then vMsg = vMsg & vbCrLf & "No Purpose Found!"
    If cboValidType.Text = "" Then vMsg = vMsg & vbCrLf & "No Validation Type Found!"
    If Not IsDate(txtDepositDate.Value) Then vMsg = "Invalid Deposit Date"
End If

    ValidateDepositSlip = vMsg
End Function

Sub SetCboListData()
On Error GoTo er

    DoEvents
    Call modObjTools.addToListObjectUsingArrayWithCols(cboBANKACCTNO, arrBANKACCT, modMotoR.AllArrBankInfo, 2, 1)

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->SetCboListData()")
End Sub

Public Sub CboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Qry = "Call sp_mrcretails.sp_MgrCshFlow_getBankAcct('" & sysBranchCode & "')"
    With mainRs
        .Open Qry, LYRcon, adOpenStatic, adLockReadOnly
        If .State = 1 Then
           Call modObjTools.addToListObjectWithArrayGivenField(cboBANKNAME, mainRs, arrBANK, "BANKID", "BANKNAME")
        End If
        If .State <> 0 Then .Close
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub cboDisplayBANKDTLS()

    If cboBANKNAME.Text <> "" Then
        'Get Data from database
        DoEvents
        Qry = "call sp_mrcretails.sp_MgrCshFlow_getBankAcctNo('" & sysBranchCode & "','" & Trim(cboBANKNAME.Text) & "')"
        With mainRs
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .State = 1 Then
                If .RecordCount > 0 Then
                    Call modObjTools.fillArray(modMotoR.AllArrBankInfo, mainRs)
                Else
                    MsgBox "No Record Found", vbInformation + vbOKOnly, "Validation"
                    Exit Sub
                End If
            End If
            If .State <> 0 Then .Close
        End With
        'Load Data to combo box
        Call SetCboListData
    End If
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    
'    cboBANKNAME.Locked = DataEntry
'    cboBANKACCTNO.Locked = DataEntry
    txtBankName.Locked = DataEntry
    txtBankAcctno.Locked = DataEntry
    txtBANKACCTNAME.Locked = DataEntry
    
    txtprpby.Locked = DataEntry
    'txttotal.Locked = DataEntry
    txtcash.Locked = DataEntry
    txtcheque.Locked = DataEntry
    'txtLess.Locked = DataEntry
    txtactcash.Locked = DataEntry
    txtdepby.Locked = DataEntry
    txtvalnum.Locked = DataEntry
    'cboValidType.Locked = DataEntry
    
    txtCVNumber.Locked = DataEntry
    txtReason.Locked = DataEntry
    txtPurpose.Locked = DataEntry
    
    txtCVDATE.Locked = True
    txtBOFAmnt.Locked = True
    
    If txtBankAcctno.Locked Then
        txtBankName.BackColor = txtHoldGrayBackColor.BackColor
        txtBankAcctno.BackColor = txtHoldGrayBackColor.BackColor
'        cboBANKNAME.BackColor = txtHoldGrayBackColor.BackColor
'        cboBANKACCTNO.BackColor = txtHoldGrayBackColor.BackColor
        txtBANKACCTNAME.BackColor = txtHoldGrayBackColor.BackColor
        txtprpby.BackColor = txtHoldGrayBackColor.BackColor
        'txttotal.BackColor = txtHoldGrayBackColor.BackColor
        txtcash.BackColor = txtHoldGrayBackColor.BackColor
        txtcheque.BackColor = txtHoldGrayBackColor.BackColor
        'txtLess.BackColor = txtHoldGrayBackColor.BackColor
        'txtDepositDate.BackColor = txtHoldGrayBackColor.BackColor
        txtactcash.BackColor = txtHoldGrayBackColor.BackColor
        txtdepby.BackColor = txtHoldGrayBackColor.BackColor
        txtvalnum.BackColor = txtHoldGrayBackColor.BackColor
        txtCVDATE.BackColor = txtHoldGrayBackColor.BackColor
        txtBOFAmnt.BackColor = txtHoldGrayBackColor.BackColor
        txtCVNumber.BackColor = txtHoldGrayBackColor.BackColor
        txtReason.BackColor = txtHoldGrayBackColor.BackColor
        txtPurpose.BackColor = txtHoldGrayBackColor.BackColor
        
        'cboValidType.BackColor = txtHoldGrayBackColor.BackColor
        txtDepositDate.Enabled = False
        'DTMgrDate.Enabled = False
        dtBOFDate.Enabled = False
        
        opt1.Enabled = False
        opt2.Enabled = False
        chkCheque.Enabled = False
    Else
'        If formAction <> "BOF" Then
'            cboBANKNAME.BackColor = txtHoldWhiteBackColor.BackColor
'            cboBANKACCTNO.BackColor = txtHoldWhiteBackColor.BackColor
'            txtBANKACCTNAME.Locked = True
'        Else
'        cboBANKNAME.Locked = True
'        cboBANKACCTNO.Locked = True
        txtBankName.Locked = True
        txtBankAcctno.Locked = True
        txtBANKACCTNAME.Locked = True
        txtprpby.Locked = True
        'txttotal.Locked = True
        txtcash.Locked = True
        txtcheque.Locked = True
        'txtLess.Locked = True
        txtDepositDate.Enabled = True
        'DTMgrDate.Enabled = True
        dtBOFDate.Enabled = True
        
        'txtDepositDate.BackColor = txtHoldWhiteBackColor.BackColor
        txtactcash.BackColor = txtHoldWhiteBackColor.BackColor
        txtdepby.BackColor = txtHoldWhiteBackColor.BackColor
        txtvalnum.BackColor = txtHoldWhiteBackColor.BackColor
        
        txtCVNumber.BackColor = txtHoldWhiteBackColor.BackColor
        txtReason.BackColor = txtHoldWhiteBackColor.BackColor
        txtPurpose.BackColor = txtHoldWhiteBackColor.BackColor
        
        'cboValidType.BackColor = txtHoldWhiteBackColor.BackColor
        opt1.Enabled = True
        opt2.Enabled = True
        chkCheque.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    ssTabNewBank.Visible = False
End Sub

Private Sub cmdOK_Click()

    If txtNewBankName.Text <> "" And txtNewBankAcctNo.Text <> "" And txtNewBankAcctName.Text <> "" Then
        Qry = "call sp_mrcretails.sp_MgrCshFlow_Bank_create('" & sysBranchCode & "','" & Trim(txtNewBankName.Text) & "', '" & Trim(txtNewBankAcctNo.Text) & "','" & Trim(txtNewBankAcctName.Text) & "'," & Val(UserID) & ")"
        With mainRs
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                If !level = "Granted" Then
                    MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                    If formAction = "add" Then
                        txtBankName.Text = txtNewBankName.Text
                        txtBankAcctno.Text = txtNewBankAcctNo.Text
                        txtBANKACCTNAME.Text = txtNewBankAcctName.Text
                        myBANKID = !BANKID
                    End If
                Else
                    MsgBox "" & !message & "", vbInformation + vbOKOnly, "" & !level & ""
                    .Close
                    Exit Sub
                End If
            End If
            .Close
        End With
    End If
    ssTabNewBank.Visible = False
    'Call CboDisplay

End Sub

Private Sub cboBANKACCTNO_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cboBANKACCTNO
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
        
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCategory-Change()")
End Sub

Private Sub cboBANKACCTNO_Click()
txtBANKACCTNAME.Text = ""
myBANKID = 0
If cboBANKNAME.Text <> "" And cboBANKACCTNO.Text <> "" Then
    txtBANKACCTNAME.Text = arrBANKACCT(cboBANKACCTNO.ListIndex, 2)
    myBANKID = arrBANKACCT(cboBANKACCTNO.ListIndex, 0)
End If
End Sub

Private Sub cboBANKACCTNO_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cboBANKACCTNO.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboBANKACCTNO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboBANKACCTNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbModel, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboBANKACCTNO_Click
    End If
End Sub

Private Sub cboBANKNAME_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cboBANKNAME
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
        
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBANKNAME-Change()")
End Sub

Private Sub cboBANKNAME_Click()
If cboBANKNAME.Text <> "" Then
'    Qry = "call sp_mrcretails.sp_MgrCshFlow_getBankAcctNo('" & sysBranchCode & "','" & Trim(cboBANKNAME.Text) & "')"
'    With mainRs
'        If .State = adStateOpen Then .Close
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            cboBANKACCTNO.Clear
'            Do While Not .EOF
'                cboBANKACCTNO.AddItem IIf(IsNull(!BANKACCTNO), "", !BANKACCTNO)
'                .MoveNext
'            Loop
'        Else
'            MsgBox "No Bank Account Found!", vbInformation + vbOKOnly, "Validation"
'        End If
'        .Close
'    End With
    Call cboDisplayBANKDTLS
Else
    MsgBox "BANK NAME IS MISSING!", vbInformation + vbOKOnly, "Validation"
    Exit Sub
End If

cboBANKACCTNO.SetFocus
End Sub

Private Sub cboBANKNAME_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboBANKNAME_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbModel, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboBANKNAME_Click
        cboBANKACCTNO.SetFocus
    End If
End Sub

Private Sub cboValidType_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboValidType
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
        
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboValidType-Change()")
End Sub

Private Sub cboValidType_Click()

    If cboValidType.Text <> "" Then
        If formAction <> "add" And cboValidType.Text <> "BOF/CASHOUT" Then
            Frame6.Visible = True 'CASH/CHECKS
            Frame5.Visible = False 'BOF
            flxChecks.Visible = True
            Frame7.Visible = False
            chkCheque.Enabled = True
    
            Qry = "call sp_mrcretails.sp_MgrCshFlow_getDepositSlip_dtls('" & sysBranchCode & "'," & Val(myMGRCFID) & ",'" & Trim(cboValidType.Text) & "')"
            With mainRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    If !level = "Granted" Then
                        .MoveFirst
                        flxChecks.Rows = .RecordCount + 1
        
'                        cboBANKNAME.Text = IIf(IsNull(!BANKNAME), "", !BANKNAME)
'                        cboBANKACCTNO.Text = IIf(IsNull(!BANKACCTNO), "", !BANKACCTNO)
                        txtBankName.Text = IIf(IsNull(!BANKNAME), "", !BANKNAME)
                        txtBankAcctno.Text = IIf(IsNull(!BANKACCTNO), "", !BANKACCTNO)
                        txtBANKACCTNAME.Text = IIf(IsNull(!BANKACCTNAME), "", !BANKACCTNAME)
                        txtDepositDate.Value = IIf(IsNull(!DATEDEPOSIT), Format(Now(), "MM/DD/YYYY"), !DATEDEPOSIT)
                        txtdepby.Text = IIf(IsNull(!DEPOSITEDBY), "", !DEPOSITEDBY)
                        txtactcash.Text = Format(IIf(IsNull(!DEPOSITAMNT), 0#, !DEPOSITAMNT), "#,###,###.00")
                        txtvalnum.Text = IIf(IsNull(!VALIDNO), "", !VALIDNO)
                        cboValidType.Text = IIf(IsNull(!VALIDTYPE), "", !VALIDTYPE)
                        myBANKID = IIf(IsNull(!BANKID), 0, !BANKID)
                        myDEPSLIPID = IIf(IsNull(!DEPSLIPID), 0, !DEPSLIPID)
                        lblValid.Caption = "VALIDATED"
                        
                        For i = 1 To flxChecks.Rows - 1
                            flxChecks.TextMatrix(i, 1) = IIf(IsNull(!CHECKNO), "", !CHECKNO)
                            flxChecks.TextMatrix(i, 2) = IIf(IsNull(!CHECKNAME), "", !CHECKNAME)
                            flxChecks.TextMatrix(i, 3) = IIf(IsNull(!CHECKTYPE), "", !CHECKTYPE)
                            flxChecks.TextMatrix(i, 4) = Format(IIf(IsNull(!Amount), "", !Amount), "#,###,###.00")
                            flxChecks.TextMatrix(i, 5) = IIf(IsNull(!CHECKDATE), "", !CHECKDATE)
                            flxChecks.TextMatrix(i, 6) = IIf(IsNull(!MGRCFID), "", !MGRCFID)
                            .MoveNext
                        Next i
                    Else
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "Validation"
                        .Close
                        Exit Sub
                    End If
                Else
                    MsgBox "No Validation Information Found! Unvalidated", vbInformation + vbOKOnly, "Validation"
                    txtBankName.Text = ""
                    txtBankAcctno.Text = ""
                    txtBANKACCTNAME.Text = ""
                    txtdepby.Text = ""
                    txtactcash.Text = 0#
                    txtvalnum.Text = ""
                    chkCheque.Value = 0
                    myBANKID = 0
                    myDEPSLIPID = 0
                    flxChecks.Rows = 1
                    flxChecks.Rows = 2
                    txtBOFAmnt.Text = 0#
                    txtCVDATE.Text = ""
                    txtCVNumber.Text = ""
                    txtReason.Text = ""
                    txtPurpose.Text = ""
                    lblValid.Caption = "UNVALIDATED"
                    formAction = ""
                    .Close
                    Exit Sub
                End If
                .Close
            End With
        ElseIf formAction = "BOF" Or cboValidType.Text = "BOF/CASHOUT" Then
            Frame6.Visible = False 'CASH/CHECKS
            Frame5.Visible = True 'BOF
            flxChecks.Visible = False
            Frame7.Visible = True
            chkCheque.Enabled = False
            
            Qry = "call sp_mrcretails.sp_MgrCshFlow_getDepositSlip_dtls('" & sysBranchCode & "'," & Val(myMGRCFID) & ",'" & Trim(cboValidType.Text) & "')"
            With mainRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    If !level = "Granted" Then
                        .MoveFirst
                        dtBOFDate.Value = IIf(IsNull(!DATEDEPOSIT), Format(Now(), "MM/DD/YYYY"), !DATEDEPOSIT)
                        txtBOFAmnt.Text = Format(IIf(IsNull(!BOFAMOUNT), 0#, !BOFAMOUNT), "#,###,###.00")
                        cboValidType.Text = IIf(IsNull(!VALIDTYPE), "", !VALIDTYPE)
                        txtCVNumber.Text = IIf(IsNull(!cvno), "", !cvno)
                        txtCVDATE.Text = IIf(IsNull(!CVDate), Format(Now(), "MM/DD/YYYY"), !CVDate)
                        txtReason.Text = IIf(IsNull(!REASON), "", !REASON)
                        txtPurpose.Text = IIf(IsNull(!PURPOSE), "", !PURPOSE)
                        myDEPSLIPID = IIf(IsNull(!DEPSLIPID), 0, !DEPSLIPID)
                        lblValid.Caption = "VALIDATED"
                    Else
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "Validation"
                        .Close
                        Exit Sub
                    End If
                Else
                    MsgBox "No Validation Information Found! Unvalidated", vbInformation + vbOKOnly, "Validation"
                    myDEPSLIPID = 0
                    formAction = ""
                    .Close
                    Exit Sub
                End If
                .Close
            End With
        End If
    End If
    
End Sub

Private Sub cboValidType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboValidType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbModel, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboValidType_Click
    End If
End Sub

Private Sub cboValidType_LostFocus()
'Call cboValidType_Click
End Sub

Private Sub chkCheque_Click()
If chkCheque.Value = 1 Then
    txtactcash.Locked = True
    txtactcash.Text = ""
    txtactcash.BackColor = txtHoldGrayBackColor.BackColor
    'lblActCash.Caption = "Check Deposit"
Else
    txtactcash.Locked = False
    txtactcash.Text = ""
    txtactcash.BackColor = txtHoldWhiteBackColor.BackColor
    flxChecks.Rows = 1
    flxChecks.Rows = 2
    'lblActCash.Caption = "Cash Deposit"
End If
End Sub

Private Sub cmdNewBank_Click()
    With ssTabNewBank
        .Top = (Me.Height - .Height) / 2
        .Left = (Me.Width - .Width) / 2
        .Visible = True
        txtNewBankName.Text = ""
        txtNewBankAcctNo.Text = ""
        txtNewBankAcctName.Text = ""
    End With
End Sub

Private Sub cmdTrans_Click(Index As Integer)
Dim countCHCKMGRCFID As Integer
Dim countDEPMGRCFID As Integer
On Error GoTo er

    If Index = 0 Then 'add
        If cboValidType <> "" Then
            If myDEPSLIPID = 0 Then
                If formAction <> "BOF" Then
                    formAction = "add"
                    Call clearform
                    Call LockData(True)
                    Call SetButtonVisibility(True)
                    cboValidType.Enabled = True
                    flxChecks.Visible = True
                    Frame7.Visible = False
                    Frame6.Visible = True 'CASH/CHECK
                    Frame5.Visible = False   'BOF
                    cboValidType.Clear
                    cboValidType.AddItem ""
                    cboValidType.AddItem "CASH"
                    cboValidType.AddItem "CHECKS"
                    chkCheque.Enabled = True
                    Call getBankDetails(myBANKID)
                    lblValid.Caption = "UNVALIDATED"
                End If
            Else
                If cboValidType.Text = "CHECKS" Then
                    countCHCKMGRCFID = Val(GetFieldValue("SELECT COUNT(CHKSPRID) FROM CHECKS WHERE MGRCFID = " & Val(myMGRCFID) & " "))
                    countDEPMGRCFID = Val(GetFieldValue("Select COUNT(CHECKSLIPINFO.DEPSLIPID) From DEPOSITSLIPINFO LEFT JOIN CHECKSLIPINFO ON CHECKSLIPINFO.DEPSLIPID = DEPOSITSLIPINFO.DEPSLIPID WHERE MGRCFID = " & Val(myMGRCFID) & " and VALIDTYPE = 'CHECKS' "))
                    If Val(countDEPMGRCFID) < Val(countCHCKMGRCFID) Then
                        formAction = "add"
                        Call clearform
                        Call LockData(True)
                        Call SetButtonVisibility(True)
                        cboValidType.Enabled = True
                        flxChecks.Visible = True
                        Frame7.Visible = False
                        Frame6.Visible = True 'CASH/CHECK
                        Frame5.Visible = False   'BOF
                        cboValidType.Clear
                        cboValidType.AddItem ""
                        cboValidType.AddItem "CASH"
                        cboValidType.AddItem "CHECKS"
                        chkCheque.Enabled = True
                        Call getBankDetails(myBANKID)
                        lblValid.Caption = "UNVALIDATED"
                    Else
                        MsgBox "It's already Validated! Please Select other Validation Type!", vbInformation + vbOKOnly, "Validation"
                        Exit Sub
                    End If
                Else
                    MsgBox "It's already Validated! Please Select other Validation Type!", vbInformation + vbOKOnly, "Validation"
                    Exit Sub
                End If
            End If
        Else
            MsgBox "Please Select Validation Type!", vbInformation + vbOKOnly, "Validation"
            Exit Sub
        End If
    ElseIf Index = 1 Then 'BOF
        formAction = "BOF"
        Call clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        Frame6.Visible = False 'CASH/CHECK
        Frame5.Visible = True   'BOF
        cboValidType.Text = "BOF/CASHOUT"
        cboValidType.Enabled = False
        flxChecks.Visible = False
        Frame7.Visible = True
        chkCheque.Enabled = False
        lblValid.Caption = "UNVALIDATED"
    ElseIf Index = 3 Then 'exit
        formAction = "EXIT"
        ProperExit = True
        Unload Me
    ElseIf Index = 4 Then 'save
        myMsg = ""
        myMsg = ValidateDepositSlip
        If myMsg <> "" Then
            MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        Ans = MsgBox("Do you want to Save Deposit Slip?", vbQuestion + vbYesNo, "Save")
        If Ans = vbNo Then Exit Sub
        myDEPSLIPID = 0
        myValidNo = ""
        myValidDate = Format(Now(), "YYYY-MM-DD")
        cboValidType.Tag = cboValidType.Text
        If formAction = "add" Or formAction = "BOF" Then
            If cboValidType.Text = "CHECKS" Then
                Call CreateTempTable
            End If
            
            If cboValidType.Text = "BOF/CASHOUT" Then
                myValidNo = Trim(txtCVNumber.Text)
                myValidDate = Format(dtBOFDate.Value, "YYYY-MM-DD")
            Else
                myValidNo = Trim(txtvalnum.Text)
                myValidDate = Format(txtDepositDate.Value, "YYYY-MM-DD")
            End If
            
            Qry = "call sp_mrcretails.sp_MgrCshFlow_create('" & sysBranchCode & "'," & Val(myBANKID) & "," & Val(myMGRCFID) & ", " & _
                  "'" & Format(myValidDate, "YYYY-MM-DD") & "','" & Trim(txtdepby.Text) & "','" & Replace(Trim(txtactcash.Text), ",", "") & "', " & _
                  "'" & Trim(myValidNo) & "','" & Trim(cboValidType.Text) & "'," & Val(UserID) & ",'" & Trim(txtReason.Text) & "','" & Trim(txtPurpose.Text) & "','" & Replace(Trim(txtBOFAmnt.Text), ",", "") & "')"
            With mainRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    If !level = "Granted" Then
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "Validation"
                        myDEPSLIPID = !DEPID
                        lblValid.Caption = "VALIDATED"
                    ElseIf !level = "Error" Then
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "Validation"
                        myDEPSLIPID = 0
                        lblValid.Caption = "UNVALIDATED"
                    End If
                End If
                .Close
            End With
        End If

        Call LockData(False)
        Call SetButtonVisibility(False)
        cboValidType.Enabled = True
        cboValidType.Clear
        cboValidType.AddItem ""
        cboValidType.AddItem "CASH"
        cboValidType.AddItem "CHECKS"
        cboValidType.AddItem "BOF/CASHOUT"
        
        cboValidType.Text = cboValidType.Tag
        formAction = ""
    ElseIf Index = 5 Then 'cancel
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        If Ans = vbNo Then Exit Sub
        Call LockData(False)
        Call SetButtonVisibility(False)
        cboValidType.Enabled = True
        cboValidType.Clear
        cboValidType.AddItem ""
        cboValidType.AddItem "CASH"
        cboValidType.AddItem "CHECKS"
        cboValidType.AddItem "BOF/CASHOUT"
        Frame6.Visible = True 'CASH/CHECK
        Frame5.Visible = False   'BOF
        formAction = ""
        flxChecks.Visible = True
        Frame7.Visible = False
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Private Sub DTMgrDate_Change()
Call ShowRecord(DTMgrDate.Value)
End Sub

Private Sub flxChecks_Click()
On Error GoTo er

    ActRow = flxChecks.Row
    ActCol = flxChecks.Col
    
    If formAction = "add" Then
        If cboValidType.Text = "CHECKS" Then
            If ActCol = 0 Or ActRow = 0 Then Exit Sub
            Select Case ActCol
                Case 1
                    With txtChckNo
                        If flxChecks.TextMatrix(ActRow, 6) <> "" Then .Tag = flxChecks.TextMatrix(ActRow, 6)
                        .Top = Frame3.Top + flxChecks.CellTop + flxChecks.Top
                        .Left = Frame3.Left + flxChecks.CellLeft + flxChecks.Left
                        .Height = flxChecks.CellHeight
                        .Width = flxChecks.CellWidth
                        .Text = flxChecks.Text
                        .Visible = True
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .ZOrder
                        .SetFocus
                    End With
            End Select
        End If
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxReplOrd-click()")
End Sub

Private Sub flxChecks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxChecks_Click
    End If
End Sub



Private Sub flxChecks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If flxChecks.TextMatrix(1, 1) <> "" Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuStatus)
        End If
    End If
End Sub

Private Sub flxCV_Click()
On Error GoTo er

    ActRow = flxCV.Row
    ActCol = flxCV.Col
    
    If formAction = "BOF" Then
        If ActCol = 0 Or ActRow = 0 Then Exit Sub
        Select Case ActCol
            Case 1
                With txtCVNO
                    If flxCV.TextMatrix(ActRow, 1) <> "" Then .Tag = flxCV.TextMatrix(ActRow, 1)
                    .Top = Frame3.Top + flxCV.CellTop + flxCV.Top
                    .Left = Frame3.Left + flxCV.CellLeft + flxCV.Left
                    .Height = flxCV.CellHeight
                    .Width = flxCV.CellWidth
                    .Text = flxCV.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
        End Select
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxReplOrd-click()")
End Sub

Private Sub Form_Load()

Call CboDisplay
txtDepositDate.Value = Format(Now(), "MM/DD/YYYY")
dtBOFDate.Value = Format(Now(), "MM/DD/YYYY")
Call LockData(False)
'===================BUTTON POSITION===================
If cmdTrans(0).Visible = False Then
    'Align Save Button
    cmdTrans(4).Left = cmdTrans(0).Left
    cmdTrans(4).Top = cmdTrans(0).Top
    'Align Cancel Button
    cmdTrans(5).Left = cmdTrans(1).Left
    cmdTrans(5).Top = cmdTrans(1).Top
End If

cboValidType.AddItem ""
cboValidType.AddItem "CASH"
cboValidType.AddItem "CHECKS"
cboValidType.AddItem "BOF/CASHOUT"

End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->PRINT
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL

    'hide all button
    For i = 0 To 3
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
    Else
    For i = 0 To 3
        cmdTrans(i).Visible = True
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    Next i
    
    End If
End Sub



Public Sub ShowRecord(myCFDate As Date)
    Dim rsCF As New ADODB.Recordset
    Dim rsBANK As New ADODB.Recordset
    Dim Qry As String, i As Integer, TtlFRec As Currency
    Dim MyBills() As String, MyCoins() As String
    
    myMGRCFID = 0
    
    Qry = "CALL sp_mrcretails.sp_MgrCshFlow_DepositValidation( '" & sysBranchCode & " ','" & Format(myCFDate, "yyyy-mm-dd") & "')"

    With rsCF
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If !level = "Error" Then
                MsgBox " " & !message & " ", vbInformation + vbOKOnly, " " & !level & ""
                txtprpby.Text = ""
                txtcash.Text = 0#
                txtcheque.Text = 0#
                'txttotal.Text = 0#
                'txtLess.Text = 0#
                myMGRCFID = 0
'                cboBANKNAME.Text = ""
'                cboBANKACCTNO.Text = ""
                txtBankName.Text = ""
                txtBankAcctno.Text = ""
                txtBANKACCTNAME.Text = ""
                txtdepby.Text = ""
                txtactcash.Text = 0#
                txtvalnum.Text = ""
                cboValidType.Text = ""
                myBANKID = 0
                myDEPSLIPID = 0
            Else
                txtprpby.Text = !PreparedBy
                txtcash.Text = Format(!CASHDEPOSIT, "#,##0.00")
                txtcheque.Text = Format(!CHECKSDEPOSIT, "#,##0.00")
                'txtLess.Text = Format(!LESS, "#,##0.00")
                'txttotal.Text = Format((!CASHDEPOSIT + !CHECKSDEPOSIT) - !LESS, "#,##0.00")
                DTMgrDate.Value = Format(!CFdate, "MM/DD/YYYY")
                myMGRCFID = !MGRCFID
'                cboBANKNAME.Text = IIf(IsNull(!BANKNAME), "", !BANKNAME)
'                cboBANKACCTNO.Text = IIf(IsNull(!BANKACCTNO), "", !BANKACCTNO)
                txtBankName.Text = IIf(IsNull(!BANKNAME), "", !BANKNAME)
                txtBankAcctno.Text = IIf(IsNull(!BANKACCTNO), "", !BANKACCTNO)
                txtBANKACCTNAME.Text = IIf(IsNull(!BANKACCTNAME), "", !BANKACCTNAME)
                txtDepositDate.Value = IIf(IsNull(!DATEDEPOSIT), Format(Now(), "MM/DD/YYYY"), !DATEDEPOSIT)
                txtdepby.Text = IIf(IsNull(!DEPOSITEDBY), "", !DEPOSITEDBY)
                txtactcash.Text = Format(IIf(IsNull(!DEPOSITAMNT), 0#, !DEPOSITAMNT), "#,###,###.00")
                txtvalnum.Text = IIf(IsNull(!VALIDNO), "", !VALIDNO)
                cboValidType.Text = IIf(IsNull(!VALIDTYPE), "", !VALIDTYPE)
                myBANKID = IIf(IsNull(!BANKID), 0, !BANKID)
                myDEPSLIPID = IIf(IsNull(!DEPSLIPID), 0, !DEPSLIPID)
                lblValid.Caption = "UNVALIDATED"
                
                txtBOFAmnt.Text = Format(IIf(IsNull(!BOFAMOUNT), 0#, !BOFAMOUNT), "#,###,###.00")
                txtCVNumber.Text = IIf(IsNull(!cvno), "", !cvno)
                txtCVDATE.Text = IIf(IsNull(!CVDate), Format(Now(), "MM/DD/YYYY"), !CVDate)
                txtReason.Text = IIf(IsNull(!REASON), "", !REASON)
                txtPurpose.Text = IIf(IsNull(!PURPOSE), "", !PURPOSE)
                dtBOFDate.Value = IIf(IsNull(!DATEDEPOSIT), Format(Now(), "MM/DD/YYYY"), !DATEDEPOSIT)
            End If
        End If
        .Close
    End With
    
    If myDEPSLIPID > 0 Then
        cboValidType.Locked = False
'        cboBANKNAME.Locked = False
'        cboBANKACCTNO.Locked = False
        cboValidType.BackColor = txtHoldWhiteBackColor.BackColor
'        cboBANKNAME.BackColor = txtHoldWhiteBackColor.BackColor
'        cboBANKACCTNO.BackColor = txtHoldWhiteBackColor.BackColor
        lblValid.Caption = "VALIDATED"
    End If
    
    If cboValidType.Text = "CHECKS" Then
        chkCheque.Value = 1
        Frame6.Visible = True 'CASH/CHECK
        Frame5.Visible = False   'BOF
        flxChecks.Visible = True
        Frame7.Visible = False
        chkCheque.Enabled = True
        Call getBankDetails(myBANKID)
    ElseIf cboValidType.Text = "CASH" Then
        Frame6.Visible = True 'CASH/CHECK
        Frame5.Visible = False   'BOF
        flxChecks.Visible = True
        Frame7.Visible = False
        chkCheque.Enabled = True
        Call getBankDetails(myBANKID)
    ElseIf cboValidType.Text = "BOF/CASHOUT" Then
        Frame6.Visible = False 'CASH/CHECK
        Frame5.Visible = True   'BOF
        flxChecks.Visible = False
        Frame7.Visible = True
        chkCheque.Enabled = False
    Else
        Frame6.Visible = True 'CASH/CHECK
        Frame5.Visible = False   'BOF
        flxChecks.Visible = True
        Frame7.Visible = False
        chkCheque.Enabled = True
        Call getBankDetails(myBANKID)
    End If
    
    If txtdepby.Text <> "" Then
        opt2.Value = True
        opt1.Value = False
    Else
        opt1.Value = True
        opt2.Value = False
    End If
    
    If txtdepby.Text <> "" Then
        opt2.Value = True
    End If
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er
    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        Call cmdTrans_Click(3)
        ProperExit = False
    Else
        Cancel = True
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->QueryUnload()")
End Sub

Private Sub mnuAdd_Click()
If formAction = "add" Then
    If flxChecks.TextMatrix(flxChecks.Rows - 1, 6) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    Else
        flxChecks.AddItem ""
        flxChecks.SetFocus
        flxChecks.Row = flxChecks.Rows - 1
        flxChecks.Col = 1
    End If
End If
End Sub

Private Sub mnuDelete_Click()
On Error GoTo er
If formAction = "add" Then
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxChecks.Rows = 2 Then
            flxChecks.Rows = 1
            flxChecks.Rows = 2
        Else
            flxChecks.RemoveItem ActRow
        End If
    End If
End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuDelete()")
End Sub



Private Sub opt1_Click()
If opt1.Value = True Then
    txtdepby.Locked = True
    txtdepby.Text = ""
    txtdepby.BackColor = txtHoldGrayBackColor.BackColor
Else
    txtdepby.Locked = False
    txtdepby.BackColor = txtHoldWhiteBackColor.BackColor
End If
End Sub

Private Sub opt2_Click()
If formAction = "add" Then
    If opt2.Value = True Then
        txtdepby.Locked = False
        txtdepby.BackColor = txtHoldWhiteBackColor.BackColor
    Else
        txtdepby.Locked = True
        txtdepby.BackColor = txtHoldGrayBackColor.BackColor
    End If
End If
End Sub

Private Sub RXResizer1_ResizeComplete()
Call flxChanger
End Sub




Private Sub txtactcash_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtModel, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        txtvalnum.SetFocus
    End If
End Sub

Private Sub txtactcash_LostFocus()
    If txtactcash.Text <> "" Then
        txtactcash.Text = Format(txtactcash.Text, "#,###,###.00")
    End If
End Sub

Private Sub txtChckNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtModel, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call txtChckNo_LostFocus
        If flxChecks.TextMatrix(ActRow, 2) <> "" Then
            Call mnuAdd_Click
            flxChecks.SetFocus
            flxChecks.Col = 1
        End If
    End If
End Sub

Public Sub ComputeTotalCheckAmount()
    intTotalAmount = 0
    With flxChecks
        For i = 1 To flxChecks.Rows - 1
            intTotalAmount = intTotalAmount + Val(Replace(.TextMatrix(i, 4), ",", ""))
        Next i
    End With
    txtactcash.Text = Format(intTotalAmount, "#,###,###.00")
End Sub

Private Sub txtChckNo_LostFocus()
    If txtChckNo.Tag = "skiplostfocus" Then
        txtChckNo.Tag = ""
        Exit Sub
    End If
    
    If txtChckNo.Text <> "" Then
        Qry = "call sp_mrcretails.sp_MgrCshFlow_getCheckDetails('" & sysBranchCode & "','" & Trim(txtChckNo.Text) & "','" & Val(myMGRCFID) & "','%" & Trim(cboBANKNAME.Text) & "%')"
        'InputBox Qry, Qry, Qry
        With mainRs
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                'CHKACCTNAME,CHECKTYPE,AMOUNT,CHECKDATE,MGRCFID
                flxChecks.TextMatrix(ActRow, 2) = !CHKACCTNAME
                flxChecks.TextMatrix(ActRow, 3) = !CHECKTYPE
                flxChecks.TextMatrix(ActRow, 4) = Format(!Amount, "#,###,###.00")
                flxChecks.TextMatrix(ActRow, 5) = !CHECKDATE
                flxChecks.TextMatrix(ActRow, 6) = !MGRCFID
            Else
                MsgBox "Unrecognized Check Number!. Please Verify you Bank Name!.", vbInformation + vbOKOnly, "Validation"
                flxChecks.TextMatrix(ActRow, 1) = ""
                flxChecks.TextMatrix(ActRow, 2) = ""
                flxChecks.TextMatrix(ActRow, 3) = ""
                flxChecks.TextMatrix(ActRow, 4) = ""
                flxChecks.TextMatrix(ActRow, 5) = ""
                flxChecks.TextMatrix(ActRow, 6) = ""
                txtChckNo.Visible = False
                .Close
                Exit Sub
            End If
            .Close
        End With
    End If
    flxChecks.TextMatrix(ActRow, 1) = txtChckNo.Text
    txtChckNo.Visible = False
    Call ComputeTotalCheckAmount
End Sub

Private Sub txtCVNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtReason.SetFocus
    End If
End Sub

Private Sub txtCVNumber_LostFocus()
Dim myValidNoExists As String
Dim myCFDate As String
Dim myDEPCFID As Long

    If txtCVNumber.Text <> "" Then
        Qry = "Select VALIDNO From DEPOSITSLIPINFO Where VALIDNO = '" & Trim(txtCVNumber.Text) & "' And VALIDTYPE = '" & Trim(cboValidType.Text) & "' "
        myValidNoExists = GetFieldValue(Qry)
        
        If myValidNoExists = "" Then
            Qry = "Select CVDATE From CASHVOUCHER WHERE VOUCHERNO = '" & Trim(txtCVNumber.Text) & "' ORDER BY CVDATE DESC LIMIT 1"
            txtCVDATE.Text = Format(GetFieldValue(Qry), "MM/DD/YYYY")
            
            If txtCVDATE.Text <> "" Then
                Qry = "Select AMOUNT From CASHVOUCHER WHERE VOUCHERNO = '" & Trim(txtCVNumber.Text) & "' ORDER BY CVDATE DESC LIMIT 1"
                txtBOFAmnt.Text = Format(GetFieldValue(Qry), "#,###,###.00")
            Else
                MsgBox "No Cash Voucher Found!", vbInformation + vbOKOnly, "Validation"
                Exit Sub
            End If
        Else
            Qry = "Select MGRCFID From DEPOSITSLIPINFO Where VALIDNO = '" & Trim(txtCVNumber.Text) & "' And VALIDTYPE = '" & Trim(cboValidType.Text) & "' "
            myDEPCFID = Val(GetFieldValue(Qry))
            
            If myDEPCFID > 0 Then
                Qry = "Select CFDATE From CASHFLOWMGR Where MGRCFID = " & Val(myDEPCFID) & " "
                myCFDate = Format(GetFieldValue(Qry), "MM/DD/YYYY")
                
                If myCFDate <> "" Then
                    MsgBox "Cash Voucher is already used in Cash Flow Date : " & Format(myCFDate, "MM/DD/YYYY") & " ", vbInformation + vbOKOnly, "Validation"
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub txtdepby_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtactcash.SetFocus
    End If
End Sub

Private Sub txtdepby_LostFocus()
    If txtdepby.Text <> "" Then
        opt2.Value = True
    End If
End Sub

Private Sub txtNewBankAcctName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cmdOK.SetFocus
    End If
End Sub

Private Sub txtNewBankAcctNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtNewBankAcctName.SetFocus
    End If
End Sub

Private Sub txtNewBankName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtNewBankAcctNo.SetFocus
    End If
End Sub

Private Sub txtPurpose_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cboBANKNAME.SetFocus
    End If
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtPurpose.SetFocus
    End If
End Sub

Private Sub txtvalnum_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cboValidType.SetFocus
    End If
End Sub
