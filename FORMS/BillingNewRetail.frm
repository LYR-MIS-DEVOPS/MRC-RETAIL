VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BillingNewRetail 
   Caption         =   "Statement of Accounts"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13890
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrLoadSummary 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4080
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   0
   End
   Begin TabDlg.SSTab TabMain 
      Height          =   8655
      Left            =   120
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   420
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   15266
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Summary"
      TabPicture(0)   =   "BillingNewRetail.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "txtCustID"
      Tab(0).Control(2)=   "cboYear"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Frame3"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Individual"
      TabPicture(1)   =   "BillingNewRetail.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblCust(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdPrint"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "flxBill"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "tabPbar"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtCustomerType"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Print Preview"
      TabPicture(2)   =   "BillingNewRetail.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "crView"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "DETAILS"
         Height          =   8175
         Left            =   -70080
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   1080
         Width           =   8655
         Begin MSFlexGridLib.MSFlexGrid flxBillDTLS 
            Height          =   7335
            Left            =   120
            TabIndex        =   11
            Tag             =   "XYF"
            Top             =   720
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   12938
            _Version        =   393216
            Cols            =   15
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"BillingNewRetail.frx":0054
         End
         Begin VB.Frame fraHoldFind 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   5280
            TabIndex        =   39
            Tag             =   "XYF"
            Top             =   120
            Width           =   3255
            Begin VB.TextBox txtFind 
               Height          =   375
               Left            =   1440
               MaxLength       =   20
               TabIndex        =   13
               Tag             =   "XYF"
               Top             =   120
               Width           =   1815
            End
            Begin lvButton.lvButtons_H cmdFind 
               Height          =   375
               Left            =   75
               TabIndex        =   12
               Tag             =   "XY"
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Caption         =   "&FIND NEXT"
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
               Image           =   "BillingNewRetail.frx":0101
               cBack           =   -2147483633
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "SUMMARY"
         Height          =   8175
         Left            =   -74880
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   1080
         Width           =   4695
         Begin MSFlexGridLib.MSFlexGrid flxBillSummary 
            Height          =   7695
            Left            =   120
            TabIndex        =   10
            Tag             =   "XYF"
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   13573
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   " |<BILLING SCHED |<CUT OFF |^COUNT |<RESPONSIBLE |<DATE CREATED |BMAINID |CUT_OFF_NF"
         End
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   -74280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   570
         Width           =   3015
      End
      Begin VB.TextBox txtCustomerType 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "xyf"
         Top             =   450
         Width           =   3855
      End
      Begin VB.TextBox txtCustID 
         Height          =   285
         Left            =   -74880
         TabIndex        =   2
         Top             =   9480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin CRVIEWERLibCtl.CRViewer crView 
         Height          =   8775
         Left            =   -74880
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   480
         Width           =   13455
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   -1  'True
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin TabDlg.SSTab tabPbar 
         Height          =   735
         Left            =   4680
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   5580
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1296
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Processing ....."
         TabPicture(0)   =   "BillingNewRetail.frx":069B
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Pbar"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin MSComctlLib.ProgressBar Pbar 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Tag             =   "xyf"
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxBill 
         Height          =   4335
         Left            =   120
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   4140
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         GridLinesFixed  =   1
         AllowUserResizing=   1
         FormatString    =   $"BillingNewRetail.frx":06B7
      End
      Begin lvButton.lvButtons_H cmdPrint 
         Height          =   375
         Left            =   12000
         TabIndex        =   34
         Tag             =   "XYF"
         Top             =   450
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Print Preview"
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
         Image           =   "BillingNewRetail.frx":079A
         cBack           =   -2147483633
      End
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   120
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   780
         Width           =   13455
         Begin MSComCtl2.DTPicker dtpCutOff 
            Height          =   375
            Left            =   9360
            TabIndex        =   52
            Tag             =   "xyf"
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            Format          =   142475265
            CurrentDate     =   -109205
         End
         Begin VB.ComboBox cboBillingSchem 
            Height          =   315
            ItemData        =   "BillingNewRetail.frx":0D34
            Left            =   9360
            List            =   "BillingNewRetail.frx":0D36
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Tag             =   "xyf"
            Top             =   960
            Width           =   2175
         End
         Begin VB.ComboBox cboMonthOfYear 
            Height          =   315
            Left            =   11640
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "xyf"
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox cboMonthOfMonth 
            Height          =   315
            ItemData        =   "BillingNewRetail.frx":0D38
            Left            =   9360
            List            =   "BillingNewRetail.frx":0D60
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Tag             =   "xyf"
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox cboBillingSched 
            Height          =   315
            ItemData        =   "BillingNewRetail.frx":0DC6
            Left            =   9360
            List            =   "BillingNewRetail.frx":0DC8
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Tag             =   "xyf"
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox cboACCOUNTTYPE 
            Height          =   315
            ItemData        =   "BillingNewRetail.frx":0DCA
            Left            =   1920
            List            =   "BillingNewRetail.frx":0DD7
            TabIndex        =   21
            Tag             =   "xyf"
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox txtPastDue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   25
            Tag             =   "xyf"
            Top             =   2520
            Width           =   2175
         End
         Begin VB.TextBox txtAddress 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   17
            Tag             =   "xyf"
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox txtAgent 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   18
            Tag             =   "xyf"
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox txtCRlimit 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   19
            Tag             =   "xyf"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtNotYetDue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   23
            Tag             =   "xyf"
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox txtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   24
            Tag             =   "xyf"
            Top             =   2160
            Width           =   2175
         End
         Begin VB.TextBox txtTotalDue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   27
            Tag             =   "xyf"
            Top             =   2880
            Width           =   2175
         End
         Begin VB.ComboBox cboCustomer 
            Height          =   315
            Left            =   1920
            TabIndex        =   15
            Tag             =   "xyf"
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtTerm 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   20
            Tag             =   "xyf"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtBillingNo 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   22
            Tag             =   "xyf"
            Top             =   240
            Width           =   2175
         End
         Begin lvButton.lvButtons_H cmdProcess 
            Height          =   330
            Left            =   5805
            TabIndex        =   16
            Tag             =   "XY"
            ToolTipText     =   "Process Billing"
            Top             =   195
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
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
            Image           =   "BillingNewRetail.frx":0DF3
            ImgSize         =   40
            cBack           =   -2147483633
         End
         Begin VB.Label Label5 
            Caption         =   "CUT - OFF"
            Height          =   255
            Left            =   7680
            TabIndex        =   51
            Tag             =   "xyf"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "BILLING SCHEME"
            Height          =   255
            Left            =   7680
            TabIndex        =   50
            Tag             =   "xyf"
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "MONTH OF"
            Height          =   255
            Left            =   7680
            TabIndex        =   49
            Tag             =   "xyf"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "SCHEDULE"
            Height          =   255
            Left            =   7680
            TabIndex        =   48
            Tag             =   "xyf"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lblCust 
            Caption         =   "TOTAL AMOUNT DUE"
            Height          =   255
            Index           =   8
            Left            =   7680
            TabIndex        =   43
            Tag             =   "xyf"
            Top             =   2910
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "NOT YET DUE"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   42
            Tag             =   "xyf"
            Top             =   2070
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "ACCOUNT TYPE"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   41
            Tag             =   "xyf"
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "PAST DUE"
            Height          =   255
            Index           =   13
            Left            =   7680
            TabIndex        =   40
            Tag             =   "xyf"
            Top             =   2550
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "CUSTOMER NAME"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   33
            Tag             =   "xyf"
            Top             =   270
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "ADDRESS"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   32
            Tag             =   "xyf"
            Top             =   630
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "CREDIT LIMIT"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   31
            Tag             =   "xyf"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "NAME OF AGENT"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   30
            Tag             =   "xyf"
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "CURRENT DUE"
            Height          =   255
            Index           =   7
            Left            =   7680
            TabIndex        =   29
            Tag             =   "xyf"
            Top             =   2190
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "TERM:"
            Height          =   255
            Index           =   10
            Left            =   3960
            TabIndex        =   28
            Tag             =   "xyf"
            Top             =   1350
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblCust 
            Caption         =   "BILLING #"
            Height          =   255
            Index           =   11
            Left            =   7680
            TabIndex        =   26
            Tag             =   "xyf"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Label Label1 
         Caption         =   "YEAR"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Tag             =   "xyf"
         Top             =   570
         Width           =   735
      End
      Begin VB.Label lblCust 
         Caption         =   "CUSTOMER TYPE"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   35
         Tag             =   "xyf"
         Top             =   450
         Width           =   1935
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   20
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "E&xit Report"
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
      Image           =   "BillingNewRetail.frx":65E5
      cBack           =   -2147483633
   End
   Begin VB.Label lblToolTip 
      Height          =   255
      Left            =   10200
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   0
      Width           =   2655
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   5355
      Top             =   -45
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14130
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuBilling 
      Caption         =   "Billing"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "BillingNewRetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAcctCust As New ADODB.Recordset
Dim rsPrn As New ADODB.Recordset
Dim CustSrch As Boolean, ThisBranch As String

Dim flxBillDTLSTTLCols As Integer
Dim flxBillDTLSRowNo As Integer
Dim flxBillDTLSColNo As Integer
Dim arrBILLINGSCHEME() As String
Dim arrBILLINGSCHEDULE() As String
Dim BillingDateSkip As String

Sub DisplayFlxSummary()
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Call flxBillSummaryColsChanger
    DoEvents
    Call modBilling.sp_billing_new_list_summary(cboYear.Text, tmpRs)
    If tmpRs.State <> 1 Then GoTo ExitGate
    If tmpRs.RecordCount <= 0 Then GoTo ExitGate
    '.FormatString = " |<BILLING SCHEDULE |<CUT OFF |^BILLING COUNT |<RESPONSIBLE |<DATE CREATED |BMAINID |CUT_OFF_NF |BILLING_SCHEDULE |BILLMO |BILLYR"
    'BMAINID, BILLING_SCHEME, CUT_OFF, TTLBILL, DATECREATED, EMPLOYEENAME, LASTUPDATE, BILLING_SCHEDULE, CUT_OFF_NF, BILLMO, BILLYR
    flxBillSummary.Rows = 1
    tmpRs.MoveFirst
    While Not tmpRs.EOF
        flxBillSummary.AddItem _
            vbTab & tmpRs!BILLING_SCHEME & _
            vbTab & tmpRs!CUT_OFF & _
            vbTab & tmpRs!TtlBill & _
            vbTab & tmpRs!EMPLOYEENAME & _
            vbTab & tmpRs!LASTUPDATE & _
            vbTab & tmpRs!BMAINID & _
            vbTab & tmpRs!CUT_OFF_NF & _
            vbTab & tmpRs!BILLING_SCHEDULE & _
            vbTab & tmpRs!BILLMO & _
            vbTab & tmpRs!BILLYR
        tmpRs.MoveNext
    Wend
ExitGate:
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    flxBillSummary.Rows = 1
    flxBillSummary.Rows = 2
    MsgBox Err.Description, vbCritical, "Error->DisplayFlxSummary()"
End Sub

Sub flxBillSummaryColsChanger()
    'BMAINID, BILLING_SCHEDULE, CUT_OFF, TTLBILL, DATECREATED, EMPLOYEENAME, LASTUPDATE, CUT_OFF_NF
    With flxBillSummary
        .Clear
        .Rows = 2
        .FixedCols = 1
        .Cols = 11
        .FormatString = " |<BILLING SCHEME |<CUT OFF |^COUNT |<RESPONSIBLE |<DATE CREATED |BMAINID |CUT_OFF_NF |BILLING_SCHEDULE |BILLMO |BILLYR"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.26 'BILLING SCHEME
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.22  'CUT OFF
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.16 'COUNT
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.45  'RESPONSIBLE
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.35 'DATE CREATED
        .ColWidth(6) = 0 'BMAINID
        .ColWidth(7) = 0 'CUT_OFF_NF
        .ColWidth(8) = 0 'BILLING_SCHEDULE
        .ColWidth(9) = 0 'BILLMO
        .ColWidth(10) = 0 'BILLYR
    End With
End Sub

Sub flxBillDTLSColsChanger()
    With flxBillDTLS
        flxBillDTLSTTLCols = 16
        .Clear
        .Rows = 2
        .FixedCols = 1
        .Cols = flxBillDTLSTTLCols
        'CUSTOMERID, CUSTOMERNAME, BILLINGNO, NOTYETDUE, CURRENT, PASTDUE, TTLBILL, ReProcess, MONTHEND, LASTUPDATE, EMPLOYEENAME, TRANSTYPE
        .FormatString = "^ |<BILLING NO |<CUSTOMER |<TRANS TYPE |>NOT YET DUE |>CURRENT |>PAST DUE |>TOTAL DUE |<LAST MODIFIED  |<RESPONSIBLE |ReProcess |CUT_OFF |BILLING_SCHEME |BILLING_SCHEDULE |BILLMO |BILLYR"
        
        .ColWidth(0) = 600
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.11 'BILLING NO
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.25 'CUSTOMER
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.13 'TRANS TYPE
        
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.12 'NOT YET DUE
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.11 'CURRENT
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.11 'PAST DUE
        
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.13 'TOTAL DUE
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.23 'LAST MODIFIED
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.36 'RESPONSIBLE
        .ColWidth(10) = 0 'ReProcess
        .ColWidth(11) = 0 'CUT_OFF
        .ColWidth(12) = 0 'BILLING_SCHEME
        .ColWidth(13) = 0 'BILLING_SCHEDULE
        .ColWidth(14) = 0 'BILLMO
        .ColWidth(15) = 0 'BILLYR
    End With
End Sub

Sub DisplayFlxBillDTLS(pBMAINID As String, BILLING_SCHEME As String, CUT_OFF_NF As String, BILLING_SCHEDULE As String, BILLMO As String, BILLYR As String)
On Error GoTo er
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Dim tmpRs As New ADODB.Recordset
    Dim FlxBillDTLSNo As Integer
    
    Call flxBillDTLSColsChanger
    DoEvents
    Call modBilling.sp_billing_new_list_bill_dtls(pBMAINID, tmpRs)
    If tmpRs.State <> 1 Then GoTo ExitGate
    If tmpRs.RecordCount <= 0 Then GoTo ExitGate
    '.FormatString = "^ |<BILLING NO |<CUSTOMER |<TRANS TYPE |>NOT YET DUE |>CURRENT |>PAST DUE |>TOTAL DUE |<LAST MODIFIED  |<RESPONSIBLE |ReProcess |CUT_OFF |BILLING_SCHEME |BILLING_SCHEDULE |BILLMO |BILLYR"
    'CUSTOMERID, CUSTOMERNAME, BILLINGNO, NOTYETDUE, CURRENT, PASTDUE, TTLBILL, ReProcess, MONTHEND, LASTUPDATE, EMPLOYEENAME, TRANSTYPE
    tmpRs.MoveFirst
    
    flxBillDTLS.Rows = 1
    FlxBillDTLSNo = 1
    While Not tmpRs.EOF
        flxBillDTLS.AddItem FlxBillDTLSNo & _
            vbTab & tmpRs!BillingNo & _
            vbTab & tmpRs!CustomerName & _
            vbTab & tmpRs!TRANSTYPE & _
            vbTab & IIf(Val(tmpRs!Notyetdue) > 0, Format(tmpRs!Notyetdue, "#,##0.00"), "") & _
            vbTab & IIf(Val(tmpRs!Current) > 0, Format(tmpRs!Current, "#,##0.00"), "") & _
            vbTab & IIf(Val(tmpRs!PASTDUE) > 0, Format(tmpRs!PASTDUE, "#,##0.00"), "") & _
            vbTab & IIf(Val(tmpRs!TtlBill) > 0, Format(tmpRs!TtlBill, "#,##0.00"), "") & _
            vbTab & tmpRs!LASTUPDATE & _
            vbTab & tmpRs!EMPLOYEENAME & _
            vbTab & tmpRs!Reprocess & _
            vbTab & tmpRs!CUT_OFF & _
            vbTab & BILLING_SCHEME & _
            vbTab & BILLING_SCHEDULE & _
            vbTab & BILLMO & _
            vbTab & BILLYR
        FlxBillDTLSNo = FlxBillDTLSNo + 1
        tmpRs.MoveNext
    Wend
ExitGate:
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    If tmpRs.State <> 0 Then tmpRs.Close
Exit Sub
er:
    flxBillDTLS.Rows = 1
    flxBillDTLS.Rows = 2
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, "Error->DisplayFlxBillDTLS()"
End Sub

Sub flxBillColsChanger()
    With flxBill
        .Clear
        .Rows = 2
        .FixedCols = 1
        
        If UCase(Trim(cboACCOUNTTYPE.Text)) = "SPARE PARTS" Then
            .Cols = 13
            '                 0  1             2              3     4         5         6           7        8        9         10        11        12
            .FormatString = "^ |^TRANS DATE |^INVOICE NO. |^TERM |^DUE DATE |^J.O. # |^REFERENCES  |>AMOUNT |>DEBIT |>CREDIT |>BALANCE |^BRN CODE |<BILL TYPE"
            .ColWidth(0) = 700
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.06 'TRANS DATE
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.07 'INVOICE NO.
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.05 'TERM
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.06 'DUE DATE
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.07 'J.O. #
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'REFERENCES
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.07 'AMOUNT
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07 'DEBIT
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.07 'CREDIT
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.07 'BALANCE
            .ColWidth(11) = (.Width - .ColWidth(0)) * 0.09 'BRN CODE
            .ColWidth(12) = (.Width - .ColWidth(0)) * 0.12 'BILL TYPE
        Else
            .Cols = 14
            '                 0  1             2              3     4         5         6           7           8       9        10       11       12         13
            .FormatString = "^ |^TRANS DATE |^INVOICE NO. |^TERM |^DUE DATE |^J.O. # |^REFERENCES |^PLATE #  |>AMOUNT |>DEBIT |>CREDIT |>BALANCE |^BRN CODE |<BILL TYPE"
            .ColWidth(0) = 700
            .ColWidth(1) = (.Width - .ColWidth(0)) * 0.06 'TRANS DATE
            .ColWidth(2) = (.Width - .ColWidth(0)) * 0.07 'INVOICE NO.
            .ColWidth(3) = (.Width - .ColWidth(0)) * 0.05 'TERM
            .ColWidth(4) = (.Width - .ColWidth(0)) * 0.06 'DUE DATE
            .ColWidth(5) = (.Width - .ColWidth(0)) * 0.07 'J.O. #
            .ColWidth(6) = (.Width - .ColWidth(0)) * 0.1 'REFERENCES
            .ColWidth(7) = (.Width - .ColWidth(0)) * 0.08 'PLATE #
            .ColWidth(8) = (.Width - .ColWidth(0)) * 0.07 'AMOUNT
            .ColWidth(9) = (.Width - .ColWidth(0)) * 0.07 'DEBIT
            .ColWidth(10) = (.Width - .ColWidth(0)) * 0.07 'CREDIT
            .ColWidth(11) = (.Width - .ColWidth(0)) * 0.07 'BALANCE
            .ColWidth(12) = (.Width - .ColWidth(0)) * 0.09 'BRN CODE
            .ColWidth(13) = (.Width - .ColWidth(0)) * 0.12 'BILL TYPE
        End If
        
    End With
End Sub

Private Sub DisplayBilling()
On Error GoTo er
    Dim CutOff As String, pBILLINGSCHEME As String, pBILLINGSCHEDULE As String
    Dim ErrMessage As String
    Dim flxRow As Integer, IndexNo As Integer
    Dim tmpRs As New ADODB.Recordset
    If Trim(cboMonthOfMonth.Text) = "" Or cboMonthOfMonth.ListIndex < 0 Then
        cboMonthOfMonth.SetFocus
        Exit Sub
    End If
    
    If Trim(cboMonthOfYear.Text) = "" Or cboMonthOfYear.ListIndex < 0 Then
        cboMonthOfYear.SetFocus
        Exit Sub
    End If
    
    If Trim(cboBillingSchem.Text) = "" Or cboBillingSchem.ListIndex < 0 Then
        cboBillingSchem.SetFocus
        Exit Sub
    End If
    
    If Trim(cboBillingSched.Text) = "" Or cboBillingSched.ListIndex < 0 Then
        cboBillingSched.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCustID.Text) = "" Then
        CustSrch = True
        cboCustomer.Text = ""
        txtCustID.Text = ""
        cboCustomer.SetFocus
        Exit Sub
    End If

    
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    tabPbar.Visible = False
    CutOff = dtpCutOff.Value
    
    pBILLINGSCHEME = arrBILLINGSCHEME(cboBillingSchem.ListIndex, 2)
    pBILLINGSCHEDULE = arrBILLINGSCHEDULE(cboBillingSched.ListIndex, 2)
    
    Call modBilling.sp_billing_new_process(Trim(txtCustID.Text), Trim(cboACCOUNTTYPE.Text) _
            , (cboMonthOfMonth.ListIndex + 1), cboMonthOfYear.Text, pBILLINGSCHEME, pBILLINGSCHEDULE, CutOff, tmpRs)
    If tmpRs.State <> 1 Then GoTo ExitGate
    If tmpRs.RecordCount <= 0 Then
        GoTo ExitGate
    End If
    
    If tmpRs!level = "Error" Then
        ErrMessage = tmpRs!message
        MsgBox ErrMessage, vbCritical, Caption
        GoTo ExitGate
        
    ElseIf tmpRs!level <> "Granted" Then
        ErrMessage = tmpRs!message
        MsgBox ErrMessage, vbExclamation, Caption
        GoTo ExitGate
    Else
        ErrMessage = tmpRs!message
    End If
    
    Call flxBillColsChanger
    
    Pbar.Value = 0
    Pbar.Max = tmpRs.RecordCount
    tabPbar.Caption = UCase("Processing " & SysBranch & ".... ")
    tabPbar.Visible = True
    tabPbar.ZOrder 0
    
    'BILLINGNO, TRANSTYPE, ACT_NYDUE, ACT_CUR, ACT_DUE, TTLDUE, refdate, INVOICE, TERM, ddate, TRA, INVAMT, PDC, pdate, RECEIPT, DEBIT, CREDIT, BALANCE, BRNCODE, BillType, Level, Message
    
    '                          0    1          2             3      4          5         6        7         8        9        10          11
    'flxBill.FormatString = "    |^INV. DATE |^INVOICE NO. |^TERM |^DUE DATE |^J.O. #  |>AMOUNT  |>DEBIT |>CREDIT |>BALANCE |^BRN CODE |<BILL TYPE"
    txtNotYetDue.Text = Format(tmpRs!ACT_NYDUE, "#,##0.00")
    txtCurrent.Text = Format(tmpRs!ACT_CUR, "#,##0.00")
    txtPastDue.Text = Format(tmpRs!ACT_DUE, "#,##0.00")
    txtTotalDue.Text = Format(tmpRs!TTLDUE, "#,##0.00")
    txtBillingNo.Text = tmpRs!BillingNo
    
    IndexNo = 1
    tmpRs.MoveFirst
    flxBill.Rows = 1
    
    
    BillingDateSkip = "Init"
    cboMonthOfMonth.ListIndex = CInt(tmpRs!BILLMO) - 1
    cboMonthOfYear.ListIndex = modTools.getListIndex(cboMonthOfYear, tmpRs!BILLYR)
    cboBillingSchem.ListIndex = modTools.getListIndex(cboBillingSchem, tmpRs!BILLING_SCHEME)
    cboBillingSched.ListIndex = modTools.getListIndex(cboBillingSched, tmpRs!BILLING_SCHEDULE)
    BillingDateSkip = ""
    Call cboBillingSchem_Click
    dtpCutOff.Value = CDate(tmpRs!CUT_OFF)
    
    While Not tmpRs.EOF
    
        If UCase(Trim(cboACCOUNTTYPE.Text)) = "SPARE PARTS" Then
            flxBill.AddItem IndexNo & _
                vbTab & tmpRs!REFDATE & _
                vbTab & tmpRs!Invoice & _
                vbTab & tmpRs!Term & _
                vbTab & tmpRs!ddate & _
                vbTab & tmpRs!TRA & _
                vbTab & tmpRs!ACCOUNTNO & _
                vbTab & IIf(Val(tmpRs!InvAmt) > 0, Format(tmpRs!InvAmt, "#,##0.00"), "") & _
                vbTab & IIf(Val(tmpRs!Debit) > 0, Format(tmpRs!Debit, "#,##0.00"), "") & _
                vbTab & IIf(Val(tmpRs!Credit) > 0, Format(tmpRs!Credit, "#,##0.00"), "") & _
                vbTab & IIf(Val(tmpRs!Balance) > 0, Format(tmpRs!Balance, "#,##0.00"), "") & _
                vbTab & tmpRs!BRNCODE & _
                vbTab & tmpRs!BillType
        Else
            flxBill.AddItem IndexNo & _
                vbTab & tmpRs!REFDATE & _
                vbTab & tmpRs!Invoice & _
                vbTab & tmpRs!Term & _
                vbTab & tmpRs!ddate & _
                vbTab & tmpRs!TRA & _
                vbTab & tmpRs!ACCOUNTNO & _
                vbTab & tmpRs!PLATENO & _
                vbTab & IIf(Val(tmpRs!InvAmt) > 0, Format(tmpRs!InvAmt, "#,##0.00"), "") & _
                vbTab & IIf(Val(tmpRs!Debit) > 0, Format(tmpRs!Debit, "#,##0.00"), "") & _
                vbTab & IIf(Val(tmpRs!Credit) > 0, Format(tmpRs!Credit, "#,##0.00"), "") & _
                vbTab & IIf(Val(tmpRs!Balance) > 0, Format(tmpRs!Balance, "#,##0.00"), "") & _
                vbTab & tmpRs!BRNCODE & _
                vbTab & tmpRs!BillType
        End If
            
        
        flxRow = flxBill.Rows - 1
        flxBill.TopRow = flxRow
        flxBill.Row = flxRow
        flxBill.RowSel = flxRow
        flxBill.Col = 0
        flxBill.ColSel = flxBill.Cols - 1
        flxBill.HighLight = flexHighlightAlways
        
        IndexNo = IndexNo + 1
        Pbar.Value = Pbar.Value + 1
        
        DoEvents
        tmpRs.MoveNext
    Wend
    
    If flxBill.Rows >= 2 Then
        flxBill.TopRow = 1
        flxBill.Row = 1
        flxBill.RowSel = 1
    End If
    
    DoEvents
    MsgBox ErrMessage, vbInformation, Caption
ExitGate:
    
    tabPbar.Visible = False
    If tmpRs.State <> 0 Then tmpRs.Close
    Screen.MousePointer = vbDefault
    Me.Enabled = True
     
Exit Sub
er:
    flxBill.Rows = 1
    flxBill.Rows = 2
    tabPbar.Visible = False
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, "Error->DisplayBilling()"
End Sub

Private Sub cboBillingSched_Click()
    If BillingDateSkip = "Init" Then Exit Sub
    If Trim(cboBillingSched.Text) = "" Then Exit Sub
    If BillingDateSkip = "Skip" Then
        BillingDateSkip = ""
        Exit Sub
    End If
    
    If Trim(cboBillingSched.Text) = "WEEKLY" Then
        If Trim(cboBillingSchem.Text) <> "WEEKLY" Then
            MsgBox "INVALID BILLING SCHEDULE!", vbExclamation, "Validation"
            cboBillingSched.ListIndex = -1
        End If
    End If
    
    Call cboBillingSchem_Click
End Sub

Private Sub cboBillingSchem_Click()
    Dim MaxDate As Date, MinDate As Date
    Dim LastDay As Integer, MonthNo As Integer, Year As Integer
    
    If BillingDateSkip = "Init" Then Exit Sub
    If Trim(cboBillingSchem.Text) = "" Then Exit Sub
    If Trim(cboMonthOfMonth.Text) = "" Then Exit Sub
    If Trim(cboMonthOfYear.Text) = "" Then Exit Sub
    
    MonthNo = (cboMonthOfMonth.ListIndex + 1)
    Year = Val(cboMonthOfYear.Text)
    LastDay = modTools.GetLastDayOfTheMonth(MonthNo, Year)
    MaxDate = CDate(DateSerial(Year, MonthNo, LastDay))
    MinDate = CDate(DateSerial(Year, MonthNo, 1))
        
    If Trim(cboBillingSchem.Text) = "MONTHLY" Then
        cboBillingSched.Locked = True
        cboBillingSched.BackColor = txtCustomerType.BackColor
        BillingDateSkip = "Skip"
        cboBillingSched.ListIndex = modTools.getListIndex(cboBillingSched, "MONTH END")
        dtpCutOff.Enabled = False
        dtpCutOff.MinDate = CDate("1601-01-01")
        dtpCutOff.MaxDate = MaxDate
        dtpCutOff.MinDate = MinDate
        dtpCutOff.Value = MaxDate
        
    ElseIf Trim(cboBillingSchem.Text) = "BI-MONTHLY" Then
        cboBillingSched.Locked = False
        cboBillingSched.BackColor = cboCustomer.BackColor
        dtpCutOff.Enabled = False
        dtpCutOff.MinDate = CDate("1601-01-01")
        dtpCutOff.MaxDate = MaxDate
        dtpCutOff.MinDate = MinDate
        If cboBillingSched.Text = "MONTH END" Then
            dtpCutOff.Value = CDate(dtpCutOff.MaxDate)
        ElseIf cboBillingSched.Text = "15TH DAY" Then
            dtpCutOff.Value = CDate(Year & "-" & MonthNo & "-15")
        Else
            cboBillingSched.ListIndex = -1
        End If
        BillingDateSkip = ""
        
    ElseIf Trim(cboBillingSchem.Text) = "WEEKLY" Then

        cboBillingSched.Locked = True
        cboBillingSched.BackColor = txtCustomerType.BackColor
        BillingDateSkip = "Skip"
        cboBillingSched.ListIndex = modTools.getListIndex(cboBillingSched, "WEEKLY")
        
        dtpCutOff.Enabled = True
        dtpCutOff.MinDate = CDate("1601-01-01")
        dtpCutOff.MaxDate = MaxDate
        dtpCutOff.MinDate = MinDate
        dtpCutOff.Value = MaxDate
    End If
    
End Sub

Private Sub cboCustomer_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    
    With cboCustomer
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsAcctCust.MoveFirst
            rsAcctCust.Move i
            txtCustID.Text = rsAcctCust!CUSTOMERID
            txtAddress.Text = rsAcctCust!address
            txtAgent.Text = IIf(IsNull(rsAcctCust!Agent), "", rsAcctCust!Agent)
            txtCRlimit.Text = Format(rsAcctCust!CRLimit, "#,##0.00")
            txtCustomerType.Text = IIf(IsNull(rsAcctCust!Ctype), "", rsAcctCust!Ctype)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustID.Text = ""
            txtAddress.Text = ""
            txtAgent.Text = ""
            txtCRlimit.Text = ""
        End If
    End With
End Sub

Private Sub cboCustomer_Click()
    Call cboCustomer_Change
End Sub

Private Sub cboCustomer_GotFocus()
   Dim Tmp As Long

   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboCustomer.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CustSrch = True
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
On Error GoTo er
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call cmdProcess_Click
        txtAddress.SetFocus
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cboCustomer_KeyPress()"
End Sub

Private Sub cboMonthOfMonth_Click()
    If BillingDateSkip = "Init" Then Exit Sub
    Call cboBillingSchem_Click
End Sub

Private Sub cboMonthOfYear_Click()
    If BillingDateSkip = "Init" Then Exit Sub
    Call cboBillingSchem_Click
End Sub

Private Sub cboYear_Click()
    If Trim(cboYear.Text) = "" Then
        cboYear.SetFocus
        Exit Sub
    End If
    
    Call flxBillDTLSColsChanger
    DoEvents
    
    Call DisplayFlxSummary
End Sub

Private Sub cmdFind_Click()
    Call modTools.flxFind(flxBillDTLS, txtFind, Me.Caption, flxBillDTLSTTLCols, flxBillDTLSRowNo, flxBillDTLSColNo)
End Sub

Private Sub cmdPrint_Click()
On Error GoTo er
    Dim crBill As New RptBilling
    Dim BillPrd As String, Branch As String
 
    If Val(txtCustID.Text) = 0 Then
        cboCustomer.Text = ""
        cboCustomer.SetFocus
        MsgBox "Please select a customer from the list!", vbExclamation + vbOKOnly, "No Record to Print"
        Exit Sub
    End If
    
    If Trim(txtBillingNo.Text) = "" Then
        MsgBox "Please process the billing record!", vbExclamation + vbOKOnly, "No Record to Print"
        Exit Sub
    End If
    
    DoEvents
    Call PreviewBilling(Trim(txtBillingNo.Text))

Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cmdPrint_Click()"
End Sub

Sub PreviewBilling(BillingNo As String)
On Error GoTo er
    Dim TRANSTYPE As String
    Dim tmpRs As New ADODB.Recordset
    TRANSTYPE = ""
    Call modBilling.sp_billing_new_preview(BillingNo, tmpRs)
    If tmpRs.State = 1 Then
        TRANSTYPE = tmpRs!TRANSTYPE
    End If
    
    If UCase(Trim(TRANSTYPE)) = "SPARE PARTS" Then
        Dim dsrSPBilling As New rptSPBilling
        Call displayCRviewer(tmpRs, crView, dsrSPBilling)
    Else
        Dim dsrSPBillingJO As New rptSPBillingJO
        Call displayCRviewer(tmpRs, crView, dsrSPBillingJO)
    End If

    Call crView.Zoom(150)
    TabMain.Tab = 2
    TabMain.TabEnabled(2) = True
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->PreviewBilling()"
End Sub

Private Sub cmdProcess_Click()
On Error GoTo er
    Call flxBillColsChanger
    txtBillingNo.Text = ""
    txtNotYetDue.Text = ""
    txtCurrent.Text = ""
    txtPastDue.Text = ""
    txtTotalDue.Text = ""
    
    If Trim(cboACCOUNTTYPE.Text) = "" Then
        MsgBox "Account type is missing!", vbExclamation, "Validation!"
        cboACCOUNTTYPE.SetFocus
        Exit Sub
    End If
    
    
    If Trim(cboCustomer.Text) = "" Or Trim(txtCustID.Text) = "" Then
        CustSrch = True
        cboCustomer.ListIndex = -1
        txtCustID.Text = ""
        cboCustomer.SetFocus
    Else
        If Val(txtCustID.Text) > 0 Then Call DisplayBilling
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cmdProcess_Click()"
End Sub

Private Sub cmdTrans_Click()
    Unload Me
    'MotoRep.imgBG.Visible = True
    MotoRep.toolBarButton.Visible = True
End Sub

Private Sub dtpCutOff_Change()
    If cboBillingSched.Text <> "WEEKLY" Then Exit Sub
    
    If Not (dtpCutOff.Month = (cboMonthOfMonth.ListIndex + 1) And dtpCutOff.Year = Val(cboMonthOfYear.Text)) Then
        dtpCutOff.Value = dtpCutOff.MaxDate
    End If
End Sub


Private Sub flxBillDTLS_DblClick()
On Error GoTo er
Dim Reprocess As String, CUT_OFF As String, BillingNo As String, CUSTOMER As String, TRANSTYPE As String
Dim BILLING_SCHEME As String, BILLING_SCHEDULE As String, BILLMO As String, BILLYR As String
'.FormatString = "^ |<BILLING NO |<CUSTOMER |<TRANS TYPE |>NOT YET DUE |>CURRENT |>PAST DUE |>TOTAL DUE |<LAST MODIFIED  |<RESPONSIBLE |ReProcess |CUT_OFF |BILLING_SCHEDULE |BILLMO |BILLYR"
'MonthName(Month(dStart)) & "-" & Year(dStart)
If flxBillDTLS.Rows > 0 Then
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    BillingNo = Trim(flxBillDTLS.TextMatrix(flxBillDTLS.Row, 1))
    CUSTOMER = flxBillDTLS.TextMatrix(flxBillDTLS.Row, 2)
    
    TRANSTYPE = flxBillDTLS.TextMatrix(flxBillDTLS.Row, 3)
    Reprocess = flxBillDTLS.TextMatrix(flxBillDTLS.Row, 10)
    CUT_OFF = flxBillDTLS.TextMatrix(flxBillDTLS.Row, 11)
    
    BILLING_SCHEME = flxBillDTLS.TextMatrix(flxBillDTLS.Row, 12)
    BILLING_SCHEDULE = flxBillDTLS.TextMatrix(flxBillDTLS.Row, 13)
    BILLMO = flxBillDTLS.TextMatrix(flxBillDTLS.Row, 14)
    BILLYR = flxBillDTLS.TextMatrix(flxBillDTLS.Row, 15)
    
    If Reprocess = "0" Then
PreviewBilling:
        If BillingNo = "" Then
            MsgBox "The billing number is not found!", vbExclamation, Caption
            flxBillDTLS.SetFocus
            Exit Sub
        Else
            Call PreviewBilling(BillingNo)
        End If
    Else
        cboACCOUNTTYPE.Text = Trim(TRANSTYPE)
        cboCustomer.Text = Trim(CUSTOMER)

        BillingDateSkip = "Init"
        cboMonthOfMonth.ListIndex = CInt(BILLMO) - 1
        cboMonthOfYear.ListIndex = modTools.getListIndex(cboMonthOfYear, BILLYR)
        cboBillingSchem.ListIndex = modTools.getListIndex(cboBillingSchem, BILLING_SCHEME)
        cboBillingSched.ListIndex = modTools.getListIndex(cboBillingSched, BILLING_SCHEDULE)
        BillingDateSkip = ""
        Call cboBillingSchem_Click
        dtpCutOff.Value = CDate(CUT_OFF)
        
        
        
        TabMain.Tab = 1
        DoEvents
        Call cmdProcess_Click
    End If
End If


Me.Enabled = True
Screen.MousePointer = vbDefault
Exit Sub
er:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "Error->flxBillDTLS_DblClick()"
End Sub

Private Sub flxBillSummary_DblClick()
On Error GoTo er
    Dim BMAINID As String, CUT_OFF_NF As String, BILLING_SCHEME As String, BILLING_SCHEDULE As String, BILLMO As String, BILLYR As String
    '                                0  1                  2         3               4             5             6         7           8                9       10
    'flxBillSummary.FormatString = "  |<BILLING SCHEME |<CUT OFF |^BILLING COUNT |<RESPONSIBLE |<DATE CREATED |BMAINID |CUT_OFF_NF |BILLING_SCHEDULE |BILLMO |BILLYR"
    If flxBillSummary.Rows > 0 Then
        BMAINID = flxBillSummary.TextMatrix(flxBillSummary.Row, 6)
        CUT_OFF_NF = flxBillSummary.TextMatrix(flxBillSummary.Row, 7)
        BILLING_SCHEME = flxBillSummary.TextMatrix(flxBillSummary.Row, 1)
        BILLING_SCHEDULE = flxBillSummary.TextMatrix(flxBillSummary.Row, 8)
        BILLMO = flxBillSummary.TextMatrix(flxBillSummary.Row, 9)
        BILLYR = flxBillSummary.TextMatrix(flxBillSummary.Row, 10)
        Call DisplayFlxBillDTLS(BMAINID, BILLING_SCHEME, CUT_OFF_NF, BILLING_SCHEDULE, BILLMO, BILLYR)
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->flxBillSummary()"
End Sub

Private Sub Form_Load()
On Error GoTo er
    Dim Qry As String
    Dim tmpRs As New ADODB.Recordset
    Dim dStart As Date, BMon As Date
    Dim i As Integer
    
    lblToolTip.ToolTipText = MotoRep.Caption
    
    TabMain.TabEnabled(2) = False 'Print Preview
    TabMain.Tab = 1
    
    Call flxBillColsChanger
    
    Call flxBillSummaryColsChanger
    
    DoEvents
    cboYear.Clear
    cboMonthOfYear.Clear
    Call modBilling.sp_billing_new_list_year(tmpRs)
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            tmpRs.MoveFirst
            Do While Not tmpRs.EOF
                Call cboYear.AddItem(tmpRs!Year)
                Call cboMonthOfYear.AddItem(tmpRs!Year)
                tmpRs.MoveNext
            Loop
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    
    'BILLING SCHEME
    'GroupNo = 24
    Qry = "Call sp_mrcretails.sp_library_list('" & sysBranchCode & "',24,'BSCHEM',0,1)"
    tmpRs.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
    Call modTools.addToListObjectWithArray(cboBillingSchem, tmpRs, arrBILLINGSCHEME, "Description2", 3)
    'arrBILLINGSCHEME(cboBillingSchem.ListIndex,0) = Description
    'arrBILLINGSCHEME(cboBillingSchem.ListIndex,1) = Description2
    'arrBILLINGSCHEME(cboBillingSchem.ListIndex,2) = LibCode
   
    'GroupNo =25
    Qry = "Call sp_mrcretails.sp_library_list('" & sysBranchCode & "',25,'BSCHED',0,1)"
    tmpRs.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
    Call modTools.addToListObjectWithArray(cboBillingSched, tmpRs, arrBILLINGSCHEDULE, "Description2", 3)
    'arrBILLINGSCHEDULE(cboBillingSched.ListIndex,0) = Description
    'arrBILLINGSCHEDULE(cboBillingSched.ListIndex,1) = Description2
    'arrBILLINGSCHEDULE(cboBillingSched.ListIndex,2) = LibCode
    
    If SysBranch = "MOTO-R REPUBLIC CORP." Then
        ThisBranch = "MRC1"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - WHOLESALE" Then
        ThisBranch = "MRC2"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN SELLING" Then
        ThisBranch = "VAN1"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN 2" Then
        ThisBranch = "VAN2"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then
        ThisBranch = "ULAS"
    ElseIf SysBranch = "MC PARTS DEPOT TORIL" Then
        ThisBranch = "TORIL"
    ElseIf SysBranch = "RUBI PANABO SPAREPARTS" Then
        ThisBranch = "PANABO"
    End If
    
    Timer1.Enabled = True
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->Form_Load()"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsAcctCust.State = adStateOpen Then rsAcctCust.Close
    If conVAN1.State = adStateOpen Then conVAN1.Close
    If conVAN2.State = adStateOpen Then conVAN2.Close
    If conMRC1.State = adStateOpen Then conMRC1.Close
End Sub


Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    With flxBill
            'flxbill:1-INV. DATE, 2-INV#, 3-Term, 4-Due Date, 5-TRA#, 6-INV.AMT, 7-PDC,8-PDCDATE,9-RECEIPT#, 10-DEBIT,
        '        11-CREDIT, 12-BALANCE, 13-BRANCH, 14-BILLTYPE, 15-SPSORDERID, 16-CHKPRID
        For i = 1 To 11
            .ColWidth(i) = .Width * 0.07
        Next i
        .ColWidth(7) = .Width * 0.1
    End With
End Sub

Private Sub Timer1_Timer()
    Dim Qry As String
    
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.AGENT, CUSTOMER.CRLIMIT, CUSTOMER.CTYPE, CUSTOMER.ADDRESS " & _
          "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
    With rsAcctCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboCustomer.Clear
            Do While Not .EOF
                cboCustomer.AddItem !CustomerName
                .MoveNext
            Loop
        End If
       
    End With
    Timer1.Enabled = False
    
    tmrLoadSummary.Enabled = True
End Sub
 
Private Sub tmrLoadSummary_Timer()
On Error GoTo er
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    tmrLoadSummary.Enabled = False
    If cboYear.ListCount > 0 Then
        cboYear.ListIndex = 0
    End If
    
    Me.Enabled = True
    Screen.MousePointer = vbDefault
Exit Sub
er:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    tmrLoadSummary.Enabled = False
    MsgBox Err.Description, vbCritical, "Error->tmrLoadSummary_Timer()"
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdFind_Click
    End If
End Sub
