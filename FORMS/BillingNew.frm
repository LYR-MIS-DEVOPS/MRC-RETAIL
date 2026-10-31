VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form BillingNew 
   Caption         =   "Statement of Accounts"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   13950
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3600
      Top             =   0
   End
   Begin TabDlg.SSTab TabMain 
      Height          =   9375
      Left            =   120
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   480
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   16536
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Individual"
      TabPicture(0)   =   "BillingNew.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCust(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPrint1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboCusType"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCustID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "flxBill"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tabPbar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Group"
      TabPicture(1)   =   "BillingNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Print Preview"
      TabPicture(2)   =   "BillingNew.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CRViewer1"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab tabPbar 
         Height          =   735
         Left            =   4680
         TabIndex        =   27
         Tag             =   "xyf"
         Top             =   5400
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1296
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Processing ....."
         TabPicture(0)   =   "BillingNew.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Pbar"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin MSComctlLib.ProgressBar Pbar 
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxBill 
         Height          =   5895
         Left            =   120
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   3120
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   10398
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         AllowUserResizing=   1
         FormatString    =   $"BillingNew.frx":0070
      End
      Begin VB.TextBox txtCustID 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   9480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboCusType 
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Tag             =   "xyf"
         Top             =   480
         Width           =   4695
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Tag             =   "xyf"
         Top             =   840
         Width           =   13455
         Begin VB.ComboBox CboClose 
            Height          =   315
            Left            =   10080
            TabIndex        =   32
            Tag             =   "xyf"
            Top             =   1800
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CheckBox ChkCloseAcct 
            Caption         =   "SHOW CLOSE ACCT."
            Height          =   255
            Left            =   8000
            TabIndex        =   31
            Tag             =   "XYF"
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox txtTerm 
            Height          =   315
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   26
            Tag             =   "xyf"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.ComboBox cboCustomer 
            Height          =   315
            Left            =   2160
            TabIndex        =   22
            Tag             =   "xyf"
            Top             =   360
            Width           =   4695
         End
         Begin VB.TextBox txtTotalDue 
            Height          =   315
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   21
            Tag             =   "xyf"
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox txtCurrent 
            Height          =   315
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   20
            Tag             =   "xyf"
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtPastDue 
            Height          =   315
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   19
            Tag             =   "xyf"
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtNotYetDue 
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   18
            Tag             =   "xyf"
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox txtCRlimit 
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   17
            Tag             =   "xyf"
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox txtAgent 
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   16
            Tag             =   "xyf"
            Top             =   1080
            Width           =   4695
         End
         Begin VB.TextBox txtAddress 
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   15
            Tag             =   "xyf"
            Top             =   720
            Width           =   4695
         End
         Begin VB.ComboBox cboBillPeriod 
            Height          =   315
            Left            =   10080
            TabIndex        =   13
            Tag             =   "xyf"
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label lblCust 
            Caption         =   "TERM:"
            Height          =   255
            Index           =   10
            Left            =   4680
            TabIndex        =   25
            Tag             =   "xyf"
            Top             =   1470
            Width           =   735
         End
         Begin VB.Label lblCust 
            Caption         =   "NOT YET DUE:"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   12
            Tag             =   "xyf"
            Top             =   1830
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "TOTAL AMOUNT DUE:"
            Height          =   255
            Index           =   8
            Left            =   8280
            TabIndex        =   11
            Tag             =   "xyf"
            Top             =   1470
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "CURRENT DUE:"
            Height          =   255
            Index           =   7
            Left            =   8280
            TabIndex        =   10
            Tag             =   "xyf"
            Top             =   1110
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "PAST DUE:"
            Height          =   255
            Index           =   6
            Left            =   8280
            TabIndex        =   9
            Tag             =   "xyf"
            Top             =   750
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "BILLING PERIOD:"
            Height          =   255
            Index           =   5
            Left            =   8280
            TabIndex        =   8
            Tag             =   "xyf"
            Top             =   390
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "NAME OF AGENT"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   7
            Tag             =   "xyf"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "CREDIT LIMIT:"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   6
            Tag             =   "xyf"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "ADDRESS:"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Tag             =   "xyf"
            Top             =   750
            Width           =   1935
         End
         Begin VB.Label lblCust 
            Caption         =   "CUSTOMER NAME:"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   4
            Tag             =   "xyf"
            Top             =   390
            Width           =   1935
         End
      End
      Begin lvButton.lvButtons_H cmdPrint1 
         Height          =   375
         Left            =   11640
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   420
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
         Image           =   "BillingNew.frx":0158
         cBack           =   -2147483633
      End
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   8775
         Left            =   -74880
         TabIndex        =   30
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
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin VB.Label lblCust 
         Caption         =   "CUSTOMER TYPE:"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   510
         Width           =   1935
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "BillingNew.frx":06F2
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   240
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14190
      m_MDI_Height    =   10515
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Menu mnuBilling 
      Caption         =   "Billing"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "BillingNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAcctCust As New ADODB.Recordset
Dim rsPrn As New ADODB.Recordset
Dim CustSrch As Boolean, ThisBranch As String
Dim DueStart As Date, DueEnd As Date


Private Sub DisplayBilling()
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String, BillType As String, MyChk As String, MyTRA As String
    Dim TtlPastDue As Currency, TtlCurrent As Currency, TtlNotDue As Currency
    Dim StartClosed As Date, DueDate As Date, TtlClose As Currency
    
    TtlPastDue = 0
    TtlCurrent = 0
    TtlNotDue = 0
    TtlClose = 0
    flxBill.Rows = 1
    'flxbill:1-INV. DATE, 2-INV#, 3-Term, 4-Due Date, 5-TRA#, 6-INV.AMT, 7-PDC,8-PDCDATE,9-RECEIPT#, 10-DEBIT,
    '        11-CREDIT, 12-BALANCE, 13-BRANCH, 14-BILLTYPE, 15-SPSORDERID, 16-CHKPRID
    Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, onaccterms.CUSTOMERID " & _
          "From onaccterms Where onaccterms.CUSTOMERID = " & Val(txtCustID.Text) & " AND onaccterms.`STATUS` = 'OPEN';"
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxBill.Rows = 1
            Pbar.Value = 0
            Pbar.Max = .RecordCount
            tabPbar.Caption = "Processing " & ThisBranch & ".... "
            tabPbar.Visible = True
            tabPbar.ZOrder 0
            Do While Not .EOF
                DueDate = DateAdd("d", !Terms, !TRANSDATE)
                Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                MyTRA = GetFieldValue(Qry)
                flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(DueDate, "mm/dd/yyyy") & vbTab & IIf(ThisBranch = "MRC2", "M2 ", "    ") & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND ledger.REFTYPE <> 'CANCELLED';"
                flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValue(Qry)), "#,##0.00")
                Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND ledger.REFTYPE <> 'CANCELLED';"
                flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
                If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                    Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                          "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                    flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValue(Qry)
                End If
                flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                flxBill.TextMatrix(flxBill.Rows - 1, 13) = ThisBranch
                If DueDate < DueStart Then
                    BillType = "Past Due"
                    TtlPastDue = TtlPastDue + !Balance
                ElseIf DueDate > DueEnd Then
                    BillType = "Not Yet Due"
                    TtlNotDue = TtlNotDue + !Balance
                Else
                    BillType = "Current"
                    TtlCurrent = TtlCurrent + !Balance
                End If
                flxBill.TextMatrix(flxBill.Rows - 1, 14) = BillType
                flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                If !CHKSPRID > 0 Then
                    Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checks Where checks.CHKSPRID = " & !CHKSPRID & ";"
                    MyChk = GetFieldValue(Qry)
                    flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                    flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                End If
                .MoveNext
                Pbar.Value = Pbar.Value + 1
                DoEvents
            Loop
            tabPbar.Visible = False
        End If
        .Close
    End With
    If ChkCloseAcct.Value = 1 Then
        StartClosed = DateAdd("M", -1 * Val(CboClose.ItemData(CboClose.ListIndex)), DueStart)
        Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, " & _
              "onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, onaccterms.CUSTOMERID, DATE_ADD(onaccterms.TRANSDATE,INTERVAL onaccterms.TERMS DAY) AS DUEDATE " & _
              "From onaccterms Where onaccterms.CUSTOMERID = " & Val(txtCustID.Text) & " AND onaccterms.`STATUS` = 'CLOSED' AND DATE_ADD(onaccterms.TRANSDATE,INTERVAL onaccterms.TERMS DAY) >= '" & Format(StartClosed, "YYYY-MM-DD") & "' AND DATE_ADD(onaccterms.TRANSDATE,INTERVAL onaccterms.TERMS DAY) <='" & Format(DueEnd, "YYYY-MM-DD") & "';"
        With rsInv
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                Pbar.Value = 0
                Pbar.Max = .RecordCount
                tabPbar.Caption = "Processing " & ThisBranch & ".... "
                tabPbar.Visible = True
                tabPbar.ZOrder 0
                Do While Not .EOF
                    Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                    MyTRA = GetFieldValue(Qry)
                    flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(!DueDate, "mm/dd/yyyy") & vbTab & IIf(ThisBranch = "MRC2", "M2  ", "     ") & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                    Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValue(Qry)), "#,##0.00")
                    Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValue(Qry)), "#,##0.00")
                    If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                        Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                              "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                        flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValue(Qry)
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                    flxBill.TextMatrix(flxBill.Rows - 1, 13) = ThisBranch
                    BillType = "Closed Account"
                    TtlClose = TtlClose + !Balance
                    flxBill.TextMatrix(flxBill.Rows - 1, 14) = BillType
                    flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                    flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                    If !CHKSPRID > 0 Then
                        Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checks Where checks.CHKSPRID = " & !CHKSPRID & ";"
                        MyChk = GetFieldValue(Qry)
                        flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                        flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                    End If
                    .MoveNext
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Loop
                tabPbar.Visible = False
            End If
            .Close
        End With
    End If
    If ThisBranch = "MRC2" Then
        If ConnectMRC1 Then
        'flxbill:1-INV. DATE, 2-INV#, 3-Term, 4-Due Date, 5-TRA#, 6-INV.AMT, 7-PDC,8-PDCDATE,9-RECEIPT#, 10-DEBIT,
        '        11-CREDIT, 12-BALANCE, 13-BRANCH, 14-BILLTYPE, 15-SPSORDERID, 16-CHKPRID
        Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID " & _
              "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) Where onaccterms.`STATUS` = 'OPEN' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "';"
        With rsInv
            .Open Qry, conMRC1, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                Pbar.Value = 0
                Pbar.Max = .RecordCount
                tabPbar.Caption = "Processing MRC1 .... "
                tabPbar.Visible = True
                tabPbar.ZOrder 0
                Do While Not .EOF
                    DueDate = DateAdd("d", !Terms, !TRANSDATE)
                    Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                    MyTRA = GetFieldValueBranch(Qry, "MRC1")
                    flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(DueDate, "mm/dd/yyyy") & vbTab & "M1  " & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                    Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                    "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValueBranch(Qry, "MRC1")), "#,##0.00")
                    Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                    "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValueBranch(Qry, "MRC1")), "#,##0.00")
                    If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                        Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                              "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                        flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValueBranch(Qry, "MRC1")
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                    flxBill.TextMatrix(flxBill.Rows - 1, 13) = "MRC1"
                    If DueDate < DueStart Then
                        BillType = "Past Due"
                        TtlPastDue = TtlPastDue + !Balance
                    ElseIf DueDate > DueEnd Then
                        BillType = "Not Yet Due"
                        TtlNotDue = TtlNotDue + !Balance
                    Else
                        BillType = "Current"
                        TtlCurrent = TtlCurrent + !Balance
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 14) = BillType
                    flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                    flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                    If !CHKSPRID > 0 Then
                        Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checkalloc INNER JOIN checks ON (checkalloc.CHKSPRID = checks.CHKSPRID) " & _
                              "Where checkalloc.CHKSPRID = " & !CHKSPRID & " AND checkalloc.BRANCH = 'MRC1'"
                        MyChk = GetFieldValue(Qry)
                        flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                        flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                    End If
                    .MoveNext
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Loop
                tabPbar.Visible = False
            End If
            .Close
        End With
        
        If ChkCloseAcct.Value = 1 Then
            StartClosed = DateAdd("M", -1 * Val(CboClose.ItemData(CboClose.ListIndex)), DueStart)
            
            Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, " & _
                  "onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) AS FIELD_1, customer.CUSTOMERID " & _
                  "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) Where onaccterms.`STATUS` = 'CLOSED' AND DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) >= '" & Format(StartClosed, "YYYY-MM-DD") & "' " & _
                  "AND DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) <= '" & Format(DueEnd, "YYYY-MM-DD") & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "';"
            With rsInv
                .Open Qry, conMRC1, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    Pbar.Value = 0
                    Pbar.Max = .RecordCount
                    tabPbar.Caption = "Processing MRC1 (CLOSED ACCOUNTS).... "
                    tabPbar.Visible = True
                    tabPbar.ZOrder 0
                    Do While Not .EOF
                        Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                        MyTRA = GetFieldValueBranch(Qry, "MRC1")
                        flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(!Field_1, "mm/dd/yyyy") & vbTab & "M1  " & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                        Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                        "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                        flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValueBranch(Qry, "MRC1")), "#,##0.00")
                        Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                        "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                        flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValueBranch(Qry, "MRC1")), "#,##0.00")
                        If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                            Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                                  "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                            flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValueBranch(Qry, "MRC1")
                        End If
                        flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                        flxBill.TextMatrix(flxBill.Rows - 1, 13) = "MRC1"
                        TtlClose = TtlClose + !Balance
                        flxBill.TextMatrix(flxBill.Rows - 1, 14) = "Closed Account"
                        flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                        flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                        If !CHKSPRID > 0 Then
                            Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checks Where checks.CHKSPRID = " & !CHKSPRID & ";"
                            MyChk = GetFieldValue(Qry)
                            flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                            flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                        End If
                        .MoveNext
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Loop
                    tabPbar.Visible = False
                End If
                .Close
            End With
        End If
        
        End If  'If ConnectMRC1 Then
        If ConnectVAN1 Then
        'flxbill:1-INV. DATE, 2-INV#, 3-Term, 4-Due Date, 5-TRA#, 6-INV.AMT, 7-PDC,8-PDCDATE,9-RECEIPT#, 10-DEBIT,
        '        11-CREDIT, 12-BALANCE, 13-BRANCH, 14-BILLTYPE, 15-SPSORDERID, 16-CHKPRID
        Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID " & _
              "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) Where onaccterms.`STATUS` = 'OPEN' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "';"
        With rsInv
            .Open Qry, conVAN1, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                Pbar.Value = 0
                Pbar.Max = .RecordCount
                tabPbar.Caption = "Processing VAN1.... "
                tabPbar.Visible = True
                tabPbar.ZOrder 0
                Do While Not .EOF
                    DueDate = DateAdd("d", !Terms, !TRANSDATE)
                    Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                    MyTRA = GetFieldValueBranch(Qry, "VAN1")
                    flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(DueDate, "mm/dd/yyyy") & vbTab & "V1  " & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                    Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                    "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValueBranch(Qry, "VAN1")), "#,##0.00")
                    Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                    "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValueBranch(Qry, "VAN1")), "#,##0.00")
                    If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                        Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                              "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                        flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValueBranch(Qry, "VAN1")
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                    flxBill.TextMatrix(flxBill.Rows - 1, 13) = "VAN1"
                    If DueDate < DueStart Then
                        BillType = "Past Due"
                        TtlPastDue = TtlPastDue + !Balance
                    ElseIf DueDate > DueEnd Then
                        BillType = "Not Yet Due"
                        TtlNotDue = TtlNotDue + !Balance
                    Else
                        BillType = "Current"
                        TtlCurrent = TtlCurrent + !Balance
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 14) = BillType
                    flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                    flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                    If !CHKSPRID > 0 Then
                        Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checkalloc INNER JOIN checks ON (checkalloc.CHKSPRID = checks.CHKSPRID) " & _
                              "Where checkalloc.CHKSPRID = " & !CHKSPRID & " AND checkalloc.BRANCH = 'VAN1'"
                        MyChk = GetFieldValue(Qry)
                        If MyChk <> "" Then
                            flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                            flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                        End If
                    End If
                    .MoveNext
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Loop
                tabPbar.Visible = False
            End If
            .Close
        End With
        
        If ChkCloseAcct.Value = 1 Then
            StartClosed = DateAdd("M", -1 * Val(CboClose.ItemData(CboClose.ListIndex)), DueStart)
            Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, " & _
                  "onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) AS FIELD_1, customer.CUSTOMERID " & _
                  "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) Where onaccterms.`STATUS` = 'CLOSED' AND DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) >= '" & Format(StartClosed, "YYYY-MM-DD") & "' " & _
                  "AND DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) <= '" & Format(DueEnd, "YYYY-MM-DD") & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "';"
            With rsInv
                .Open Qry, conVAN1, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    Pbar.Value = 0
                    Pbar.Max = .RecordCount
                    tabPbar.Caption = "Processing VAN1 (Closed Account).... "
                    tabPbar.Visible = True
                    tabPbar.ZOrder 0
                    Do While Not .EOF
                        Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                        MyTRA = GetFieldValueBranch(Qry, "VAN1")
                        flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(!Field_1, "mm/dd/yyyy") & vbTab & "V1  " & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                        Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                        "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                        flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValueBranch(Qry, "VAN1")), "#,##0.00")
                        Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                        "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                        flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValueBranch(Qry, "VAN1")), "#,##0.00")
                        If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                            Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                                  "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                            flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValueBranch(Qry, "VAN1")
                        End If
                        flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                        flxBill.TextMatrix(flxBill.Rows - 1, 13) = "VAN1"
                        TtlClose = TtlClose + !Balance
                        flxBill.TextMatrix(flxBill.Rows - 1, 14) = "Closed Account"
                        flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                        flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                        If !CHKSPRID > 0 Then
                            Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checks Where checks.CHKSPRID = " & !CHKSPRID & ";"
                            MyChk = GetFieldValue(Qry)
                            flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                            flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                        End If
                        .MoveNext
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Loop
                    tabPbar.Visible = False
                End If
                .Close
            End With
        End If
        
        End If  'If ConnectVAN1 Then
        
        If ConnectVAN2 Then
        'flxbill:1-INV. DATE, 2-INV#, 3-Term, 4-Due Date, 5-TRA#, 6-INV.AMT, 7-PDC,8-PDCDATE,9-RECEIPT#, 10-DEBIT,
        '        11-CREDIT, 12-BALANCE, 13-BRANCH, 14-BILLTYPE, 15-SPSORDERID, 16-CHKPRID
          Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID " & _
              "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) Where onaccterms.`STATUS` = 'OPEN' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "';"
          With rsInv
            .Open Qry, conVAN2, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                Pbar.Value = 0
                Pbar.Max = .RecordCount
                tabPbar.Caption = "Processing VAN2.... "
                tabPbar.Visible = True
                tabPbar.ZOrder 0
                Do While Not .EOF
                    DueDate = DateAdd("d", !Terms, !TRANSDATE)
                    Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                    MyTRA = GetFieldValueBranch(Qry, "VAN2")
                    flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(DueDate, "mm/dd/yyyy") & vbTab & "V2  " & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                    Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                    "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValueBranch(Qry, "VAN2")), "#,##0.00")
                    Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                    "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValueBranch(Qry, "VAN2")), "#,##0.00")
                    If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                        Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                              "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                        flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValueBranch(Qry, "VAN2")
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                    flxBill.TextMatrix(flxBill.Rows - 1, 13) = "VAN2"
                    If DueDate < DueStart Then
                        BillType = "Past Due"
                        TtlPastDue = TtlPastDue + !Balance
                    ElseIf DueDate > DueEnd Then
                        BillType = "Not Yet Due"
                        TtlNotDue = TtlNotDue + !Balance
                    Else
                        BillType = "Current"
                        TtlCurrent = TtlCurrent + !Balance
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 14) = BillType
                    flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                    flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                    If !CHKSPRID > 0 Then
                        Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checkalloc INNER JOIN checks ON (checkalloc.CHKSPRID = checks.CHKSPRID) " & _
                              "Where checkalloc.CHKSPRID = " & !CHKSPRID & " AND checkalloc.BRANCH = 'VAN2'"
                        MyChk = GetFieldValue(Qry)
                        flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                        flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                    End If
                    .MoveNext
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Loop
                tabPbar.Visible = False
            End If
            .Close
          End With
          If ChkCloseAcct.Value = 1 Then
            StartClosed = DateAdd("M", -1 * Val(CboClose.ItemData(CboClose.ListIndex)), DueStart)
            Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, " & _
                  "onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) AS FIELD_1, customer.CUSTOMERID " & _
                  "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) Where onaccterms.`STATUS` = 'CLOSED' AND DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) >= '" & Format(StartClosed, "YYYY-MM-DD") & "' " & _
                  "AND DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) <= '" & Format(DueEnd, "YYYY-MM-DD") & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "';"
            With rsInv
                .Open Qry, conVAN2, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    Pbar.Value = 0
                    Pbar.Max = .RecordCount
                    tabPbar.Caption = "Processing VAN2 (Closed Account).... "
                    tabPbar.Visible = True
                    tabPbar.ZOrder 0
                    Do While Not .EOF
                        Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                        MyTRA = GetFieldValueBranch(Qry, "VAN2")
                        flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(!Field_1, "mm/dd/yyyy") & vbTab & "V2  " & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                        Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                        "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                        flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValueBranch(Qry, "VAN2")), "#,##0.00")
                        Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                        "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                        flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValueBranch(Qry, "VAN2")), "#,##0.00")
                        If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                            Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                                  "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                            flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValueBranch(Qry, "VAN2")
                        End If
                        flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                        flxBill.TextMatrix(flxBill.Rows - 1, 13) = "VAN2"
                        TtlClose = TtlClose + !Balance
                        flxBill.TextMatrix(flxBill.Rows - 1, 14) = "Closed Account"
                        flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                        flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                        If !CHKSPRID > 0 Then
                            Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checks Where checks.CHKSPRID = " & !CHKSPRID & ";"
                            MyChk = GetFieldValue(Qry)
                            flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                            flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                        End If
                        .MoveNext
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Loop
                    tabPbar.Visible = False
                End If
                .Close
            End With
          End If
          
        End If
   
        If ConnectVan3 Then
        'flxbill:1-INV. DATE, 2-INV#, 3-Term, 4-Due Date, 5-TRA#, 6-INV.AMT, 7-PDC,8-PDCDATE,9-RECEIPT#, 10-DEBIT,
        '        11-CREDIT, 12-BALANCE, 13-BRANCH, 14-BILLTYPE, 15-SPSORDERID, 16-CHKPRID
          Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, customer.CUSTOMERNAME, customer.CUSTOMERID " & _
              "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) Where onaccterms.`STATUS` = 'OPEN' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "';"
          With rsInv
            .Open Qry, conVAN3, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                Pbar.Value = 0
                Pbar.Max = .RecordCount
                tabPbar.Caption = "Processing VAN3.... "
                tabPbar.Visible = True
                tabPbar.ZOrder 0
                Do While Not .EOF
                    DueDate = DateAdd("d", !Terms, !TRANSDATE)
                    Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                    MyTRA = GetFieldValueBranch(Qry, "VAN3")
                    flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(DueDate, "mm/dd/yyyy") & vbTab & "V3  " & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                    Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                    "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValueBranch(Qry, "VAN3")), "#,##0.00")
                    Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                    "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                    flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValueBranch(Qry, "VAN3")), "#,##0.00")
                    If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                        Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                              "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                        flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValueBranch(Qry, "VAN3")
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                    flxBill.TextMatrix(flxBill.Rows - 1, 13) = "VAN3"
                    If DueDate < DueStart Then
                        BillType = "Past Due"
                        TtlPastDue = TtlPastDue + !Balance
                    ElseIf DueDate > DueEnd Then
                        BillType = "Not Yet Due"
                        TtlNotDue = TtlNotDue + !Balance
                    Else
                        BillType = "Current"
                        TtlCurrent = TtlCurrent + !Balance
                    End If
                    flxBill.TextMatrix(flxBill.Rows - 1, 14) = BillType
                    flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                    flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                    If !CHKSPRID > 0 Then
                        Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checkalloc INNER JOIN checks ON (checkalloc.CHKSPRID = checks.CHKSPRID) " & _
                              "Where checkalloc.CHKSPRID = " & !CHKSPRID & " AND checkalloc.BRANCH = 'VAN3'"
                        MyChk = GetFieldValue(Qry)
                        flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                        flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                    End If
                    .MoveNext
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Loop
                tabPbar.Visible = False
            End If
            .Close
          End With
          If ChkCloseAcct.Value = 1 Then
            StartClosed = DateAdd("M", -1 * Val(CboClose.ItemData(CboClose.ListIndex)), DueStart)
            Qry = "SELECT onaccterms.TRANSDATE, onaccterms.SPINVNO, onaccterms.ACCTAMOUNT, onaccterms.CHKSPRID, onaccterms.BALANCE, " & _
                  "onaccterms.TERMS, onaccterms.TTLAMTPAID, onaccterms.SPSORDERID, DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) AS FIELD_1, customer.CUSTOMERID " & _
                  "From onaccterms INNER JOIN customer ON (onaccterms.CUSTOMERID = customer.CUSTOMERID) Where onaccterms.`STATUS` = 'CLOSED' AND DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) >= '" & Format(StartClosed, "YYYY-MM-DD") & "' " & _
                  "AND DATE_ADD(onaccterms.TRANSDATE, INTERVAL onaccterms.TERMS DAY) <= '" & Format(DueEnd, "YYYY-MM-DD") & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "';"
            With rsInv
                .Open Qry, conVAN3, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    Pbar.Value = 0
                    Pbar.Max = .RecordCount
                    tabPbar.Caption = "Processing VAN3 (Closed Account).... "
                    tabPbar.Visible = True
                    tabPbar.ZOrder 0
                    Do While Not .EOF
                        Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.SPSORDERID = " & !SPSORDERID & ";"
                        MyTRA = GetFieldValueBranch(Qry, "VAN3")
                        flxBill.AddItem flxBill.Rows & vbTab & Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & !Terms & vbTab & Format(!Field_1, "mm/dd/yyyy") & vbTab & "V3  " & MyTRA & vbTab & Format(!AcctAmount, "#,##0.00")
                        Qry = "SELECT SUM(Ledger.DEBIT) As TtlDebit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                        "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                        flxBill.TextMatrix(flxBill.Rows - 1, 10) = Format(!AcctAmount + Val(GetFieldValueBranch(Qry, "VAN3")), "#,##0.00")
                        Qry = "SELECT SUM(Ledger.CREDIT) As TtlCredit FROM Ledger INNER JOIN customer ON (Ledger.CUSTOMERID = customer.CUSTOMERID) " & _
                        "Where ledger.SPINVNO = '" & !SPINVNO & "' AND customer.CUSTOMERNAME = '" & cboCustomer.Text & "' AND ledger.REFTYPE <> 'CANCELLED';"
                        flxBill.TextMatrix(flxBill.Rows - 1, 11) = Format(Val(GetFieldValueBranch(Qry, "VAN3")), "#,##0.00")
                        If Val(flxBill.TextMatrix(flxBill.Rows - 1, 11)) > 0 Then
                            Qry = "SELECT Ledger.RefNo From Ledger Where ledger.SPINVNO = '" & !SPINVNO & "' AND ledger.CUSTOMERID = " & Val(txtCustID.Text) & " AND " & _
                                  "ledger.REFTYPE = 'Payment' Order By Ledger.SPINVNO Desc;"
                            flxBill.TextMatrix(flxBill.Rows - 1, 9) = GetFieldValueBranch(Qry, "VAN3")
                        End If
                        flxBill.TextMatrix(flxBill.Rows - 1, 12) = Format(!Balance, "#,##0.00")
                        flxBill.TextMatrix(flxBill.Rows - 1, 13) = "VAN3"
                        TtlClose = TtlClose + !Balance
                        flxBill.TextMatrix(flxBill.Rows - 1, 14) = "Closed Account"
                        flxBill.TextMatrix(flxBill.Rows - 1, 15) = !SPSORDERID
                        flxBill.TextMatrix(flxBill.Rows - 1, 16) = !CHKSPRID
                        If !CHKSPRID > 0 Then
                            Qry = "SELECT concat(checks.CHECKNO,'|',checks.CHECKDATE) as MyChk From checks Where checks.CHKSPRID = " & !CHKSPRID & ";"
                            MyChk = GetFieldValue(Qry)
                            flxBill.TextMatrix(flxBill.Rows - 1, 7) = Left$(MyChk, InStr(MyChk, "|") - 1)
                            flxBill.TextMatrix(flxBill.Rows - 1, 8) = Right$(MyChk, 10)
                        End If
                        .MoveNext
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Loop
                    tabPbar.Visible = False
                End If
                .Close
            End With
          End If
          
        End If
    
    End If  'If ThisBranch = "MRC2" Then
              
    txtPastDue.Text = Format(TtlPastDue, "#,##0.00")
    txtCurrent.Text = Format(TtlCurrent, "#,##0.00")
    txtTotalDue.Text = Format(TtlPastDue + TtlCurrent, "#,##0.00")
    txtNotYetDue.Text = Format(TtlNotDue, "#,##0.00")
    
    
    
    
End Sub

Private Sub cboBillPeriod_Click()
    DueStart = DateValue(Left$(cboBillPeriod.Text, Len(cboBillPeriod.Text) - 5) & " 1 " & Right$(cboBillPeriod.Text, 4))
    DueEnd = DateAdd("d", -1, DateAdd("m", 1, DueStart))
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
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsAcctCust.MoveFirst
            rsAcctCust.Move i
            txtCustID.Text = rsAcctCust!CUSTOMERID
            txtAddress.Text = rsAcctCust!Address
            txtAgent.Text = IIf(IsNull(rsAcctCust!Agent), "", rsAcctCust!Agent)
            txtCRlimit.Text = Format(rsAcctCust!CRLimit, "#,##0.00")
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
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboCustomer.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CustSrch = True
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAddress.SetFocus
End Sub

Private Sub cboCustomer_LostFocus()
    If Val(txtCustID.Text) > 0 Then Call DisplayBilling
End Sub




Private Sub cboCusType_Change()
'    Timer1.Enabled = True
'    Call Timer1_Timer
   ' Call custType
End Sub

Private Sub cboCusType_Click()
    If rsAcctCust.State <> 0 Then rsAcctCust.Close
End Sub

Private Sub cboCusType_LostFocus()
    Call custType
End Sub

Private Sub ChkCloseAcct_Click()
    If ChkCloseAcct.Value = 1 Then
        With CboClose
            .Clear
            .AddItem "1 Month Before Bill"
            .ItemData(.ListCount - 1) = 1
            .AddItem "2 Months Before Bill"
            .ItemData(.ListCount - 1) = 2
            .AddItem "3 Month Before Bill"
            .ItemData(.ListCount - 1) = 3
            .AddItem "4 Month Before Bill"
            .ItemData(.ListCount - 1) = 4
            .AddItem "5 Month Before Bill"
            .ItemData(.ListCount - 1) = 5
            .Visible = True
        End With
    Else
        CboClose.Visible = False
    End If
End Sub

Private Sub cmdPrint1_Click()
    Dim crBill As New RptBilling
    Dim BillPrd As String, Branch As String
    Dim i As Integer
        
    If Val(txtCustID.Text) = 0 Then
        MsgBox "Please select a customer from the list!", vbExclamation + vbOKOnly, "No Record to Print"
        Exit Sub
    End If
    
    LYRcon.Execute "DELETE printbill.* From printbill"
    
    Pbar.Value = 0
    Pbar.Max = flxBill.Rows - 1
    tabPbar.Caption = "Processing Print Preview .... "
    tabPbar.Visible = True
    tabPbar.ZOrder 0
    
    
    BillPrd = MonthName(Month(DueStart)) & " " & Day(DueStart) & " to " & Day(DueEnd) & ", " & Year(DueStart)
    
    With flxBill
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 8) = "" Then
                LYRcon.Execute "INSERT INTO printbill(CUSTOMER, ADDRESS, BILLPRD, PASTDUE, CURRENT, TTLDUE, NOTYETDUE, INVDATE, INVOICE, TERM, DUEDATE, TRA, INVAMT, PDC, PDCDATE, RECEIPT, DEBIT, CREDIT, BALANCE, BRNCODE, BILLTYPE) " & _
                "VALUES('" & cboCustomer.Text & "', '" & txtAddress.Text & "', '" & BillPrd & "', '" & Val(Replace(txtPastDue.Text, ",", "")) & "', '" & Val(Replace(txtCurrent.Text, ",", "")) & "', '" & Val(Replace(txtTotalDue.Text, ",", "")) & "', " & _
                "'" & Val(Replace(txtNotYetDue.Text, ",", "")) & "', '" & Format(DateValue(.TextMatrix(i, 1)), "yyyy-mm-dd") & "', '" & .TextMatrix(i, 2) & "', " & Val(.TextMatrix(i, 3)) & ", '" & Format(DateValue(.TextMatrix(i, 4)), "yyyy-mm-dd") & "', " & _
                "'" & .TextMatrix(i, 5) & "', '" & Val(Replace(.TextMatrix(i, 6), ",", "")) & "', '" & .TextMatrix(i, 7) & "', NULL, " & _
                "'" & .TextMatrix(i, 9) & "', '" & Val(Replace(.TextMatrix(i, 10), ",", "")) & "', '" & Val(Replace(.TextMatrix(i, 11), ",", "")) & "', '" & Val(Replace(.TextMatrix(i, 12), ",", "")) & "', '" & .TextMatrix(i, 13) & "', '" & .TextMatrix(i, 14) & "')"
            Else
                LYRcon.Execute "INSERT INTO printbill(CUSTOMER, ADDRESS, BILLPRD, PASTDUE, CURRENT, TTLDUE, NOTYETDUE, INVDATE, INVOICE, TERM, DUEDATE, TRA, INVAMT, PDC, PDCDATE, RECEIPT, DEBIT, CREDIT, BALANCE, BRNCODE, BILLTYPE) " & _
                "VALUES('" & cboCustomer.Text & "', '" & txtAddress.Text & "', '" & BillPrd & "', '" & Val(Replace(txtPastDue.Text, ",", "")) & "', '" & Val(Replace(txtCurrent.Text, ",", "")) & "', '" & Val(Replace(txtTotalDue.Text, ",", "")) & "', " & _
                "'" & Val(Replace(txtNotYetDue.Text, ",", "")) & "', '" & Format(DateValue(.TextMatrix(i, 1)), "yyyy-mm-dd") & "', '" & .TextMatrix(i, 2) & "', " & Val(.TextMatrix(i, 3)) & ", '" & Format(DateValue(.TextMatrix(i, 4)), "yyyy-mm-dd") & "', " & _
                "'" & .TextMatrix(i, 5) & "', '" & Val(Replace(.TextMatrix(i, 6), ",", "")) & "', '" & .TextMatrix(i, 7) & "', '" & Format(DateValue(.TextMatrix(i, 8)), "yyyy-mm-dd") & "', " & _
                "'" & .TextMatrix(i, 9) & "', '" & Val(Replace(.TextMatrix(i, 10), ",", "")) & "', '" & Val(Replace(.TextMatrix(i, 11), ",", "")) & "', '" & Val(Replace(.TextMatrix(i, 12), ",", "")) & "', '" & .TextMatrix(i, 13) & "', '" & .TextMatrix(i, 14) & "')"
            End If
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    tabPbar.Visible = False
        
    Qry = "SELECT printbill.CUSTOMER, printbill.ADDRESS, printbill.BILLPRD, printbill.PASTDUE, printbill.CURRENT, printbill.TTLDUE, printbill.NOTYETDUE, printbill.INVDATE, printbill.INVOICE, printbill.TERM, printbill.DUEDATE, printbill.TRA, printbill.INVAMT, printbill.PDC, " & _
          "printbill.PDCDATE, printbill.RECEIPT, printbill.DEBIT, printbill.CREDIT, printbill.BALANCE, printbill.BRNCODE, printbill.BillType From printbill Order By printbill.BILLTYPE, printbill.InvDate;"
    If SysBranch = "MC PARTS DEPOT CORP." Then
        Branch = "MC PARTS DEPOT CORPORATION"
    Else
        Branch = "MOTO-R REPUBLIC CORPORATION"
    End If
    If rsPrn.State = adStateOpen Then rsPrn.Close
    rsPrn.Open Qry, LYRcon, adOpenDynamic, adLockOptimistic
    'crSOA.DiscardSavedData
    crBill.Database.SetDataSource rsPrn
    crBill.ParameterFields(1).ClearCurrentValueAndRange
    crBill.ParameterFields(1).AddCurrentValue Branch
    crBill.ParameterFields(2).ClearCurrentValueAndRange
    crBill.ParameterFields(2).AddCurrentValue SysAddress
'    crSOA.ParameterFields(3).ClearCurrentValueAndRange
'    crSOA.ParameterFields(3).AddCurrentValue GetSettings("CASHIER")
'    crSOA.ParameterFields(4).ClearCurrentValueAndRange
'    crSOA.ParameterFields(4).AddCurrentValue GetSettings("MANAGER")
'    crSOA.ParameterFields(5).ClearCurrentValueAndRange
'    crSOA.ParameterFields(5).AddCurrentValue GetSettings("OIC")
    CRViewer1.ReportSource = crBill
    CRViewer1.ViewReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.EnableExportButton = True
    TabMain.Tab = 2
    TabMain.TabEnabled(2) = True
    
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Unload Me
    MotoRep.tbMain.Visible = True
End Sub

Private Sub Form_Load()
    Dim dStart As Date, BMon As Date
    Dim i As Integer
    
    TabMain.TabEnabled(1) = False
    TabMain.TabEnabled(2) = False
    TabMain.Tab = 0
    
'    cboCusType.AddItem "WHOLESALE"
'    cboCusType.ListIndex = 0
    
    '------ MRC RETAIL BILLING
    Qry = "SELECT DISTINCT customer.CTYPE From customer Group By CTYPE"
        
    With cboClass
        If rsAnyUse.State <> 0 Then rsAnyUse.Close
        Set rsAnyUse = LYRcon.Execute(Qry)
            If rsAnyUse.RecordCount > 0 Then
                rsAnyUse.MoveFirst
                Do While Not rsAnyUse.EOF
                    cboCusType.AddItem rsAnyUse!Ctype
                rsAnyUse.MoveNext
                Loop
            End If
        rsAnyUse.Close
    End With
    cboCusType.Text = "INDIVIDUAL"
    '------ MRC RETAIL BILLING
    
    dStart = DateAdd("M", 2, Date)
    cboBillPeriod.AddItem MonthName(Month(dStart)) & "-" & Year(dStart)
    For i = 1 To 12
        BMon = DateAdd("M", -i, dStart)
        cboBillPeriod.AddItem MonthName(Month(BMon)) & "-" & Year(BMon)
    Next i
    cboBillPeriod.ListIndex = 2
    
    If SysBranch = "MOTO-R REPUBLIC CORP." Then
        ThisBranch = "MRC1"
    ElseIf SysBranch = "MRC 2 WHOLESALE" Then
        ThisBranch = "MRC2"
    ElseIf SysBranch = "MOTO-R VAN SELLING" Then
        ThisBranch = "VAN1"
    ElseIf SysBranch = "MRC VAN 2" Then
        ThisBranch = "VAN2"
    ElseIf SysBranch = "MOTO-R ULAS" Then
        ThisBranch = "ULAS"
    ElseIf SysBranch = "MC PARTS DEPOT CORP." Then
        ThisBranch = "TORIL"
    ElseIf SysBranch = "RUBI PANABO SPAREPARTS" Then
        ThisBranch = "PANABO"
    End If
    
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
        For i = 1 To 13
            .ColWidth(i) = .Width * 0.07
        Next i
        .ColWidth(7) = .Width * 0.1
        .ColWidth(14) = .Width * 0.1
        .ColWidth(15) = 0
        .ColWidth(16) = 0
    End With
End Sub

Private Sub Timer1_Timer()
    'Call custType
    Timer1.Enabled = False
End Sub

Public Sub custType()

    Dim Qry As String
    
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.AGENT, CUSTOMER.CRLIMIT, CUSTOMER.CTYPE, CUSTOMER.ADDRESS " & _
          "From CUSTOMER Where CUSTOMER.CTYPE = '" & cboCusType.Text & "' ORDER BY CUSTOMER.CUSTOMERNAME;"
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
    
End Sub

