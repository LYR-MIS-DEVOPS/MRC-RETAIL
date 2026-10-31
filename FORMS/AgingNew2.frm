VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form AgingNew 
   Caption         =   "Aging of Accounts Receivable"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSQLsum 
      Height          =   495
      Left            =   16080
      TabIndex        =   23
      Top             =   7920
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.ComboBox cboAgingPeriod 
      Height          =   315
      Left            =   12240
      TabIndex        =   21
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtSQL 
      Height          =   495
      Left            =   16080
      TabIndex        =   4
      Top             =   7320
      Visible         =   0   'False
      Width           =   5415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAccts 
      Height          =   5055
      Left            =   16080
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   8916
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab tabAging 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   480
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   15690
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Aging Summary"
      TabPicture(0)   =   "AgingNew2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxAgeSum"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Aging Details"
      TabPicture(1)   =   "AgingNew2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "FlxAging"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Print Preview"
      TabPicture(2)   =   "AgingNew2.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "CRViewer1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   -74880
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   360
         Width           =   14385
         Begin VB.ComboBox cboCTypSum 
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
            Left            =   1680
            TabIndex        =   16
            Tag             =   "XYF"
            Top             =   172
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   12000
            TabIndex        =   15
            Tag             =   "XYF"
            Top             =   195
            Width           =   1935
         End
         Begin lvButton.lvButtons_H lvButtons_H1 
            Height          =   375
            Left            =   10680
            TabIndex        =   17
            Tag             =   "XY"
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Find Next"
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
            Image           =   "AgingNew2.frx":0054
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdPrnSum 
            Height          =   375
            Left            =   4920
            TabIndex        =   18
            Tag             =   "XY"
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
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
            Image           =   "AgingNew2.frx":05EE
            cBack           =   -2147483633
         End
         Begin VB.Label Label1 
            Caption         =   "Customer Type:"
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
            Left            =   120
            TabIndex        =   19
            Top             =   210
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   -74880
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   360
         Width           =   14385
         Begin VB.TextBox txtFind 
            Height          =   285
            Left            =   12000
            TabIndex        =   8
            Tag             =   "XYF"
            Top             =   195
            Width           =   1935
         End
         Begin VB.ComboBox cboCType 
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
            Left            =   1680
            TabIndex        =   7
            Tag             =   "XYF"
            Top             =   172
            Width           =   3135
         End
         Begin lvButton.lvButtons_H cmdFindC 
            Height          =   375
            Left            =   10680
            TabIndex        =   9
            Tag             =   "XY"
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Find Next"
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
            Image           =   "AgingNew2.frx":0B88
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCReprocess 
            Height          =   375
            Left            =   6480
            TabIndex        =   10
            Tag             =   "XY"
            Top             =   150
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "&Reprocess"
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
            Image           =   "AgingNew2.frx":1122
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdPrnDtls 
            Height          =   375
            Left            =   4920
            TabIndex        =   11
            Tag             =   "XY"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
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
            Image           =   "AgingNew2.frx":16BC
            cBack           =   -2147483633
         End
         Begin VB.Label Label2 
            Caption         =   "Customer Type:"
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
            Left            =   120
            TabIndex        =   12
            Top             =   210
            Width           =   1575
         End
      End
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   8295
         Left            =   120
         TabIndex        =   5
         Tag             =   "xyf"
         Top             =   480
         Width           =   14415
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
      Begin MSFlexGridLib.MSFlexGrid FlxAging 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   960
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13785
         _Version        =   393216
         Cols            =   24
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         FormatString    =   $"AgingNew2.frx":1C56
      End
      Begin MSFlexGridLib.MSFlexGrid flxAgeSum 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   22
         Tag             =   "xyf"
         Top             =   960
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13785
         _Version        =   393216
         Cols            =   24
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         FormatString    =   $"AgingNew2.frx":1D78
      End
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   9360
      Visible         =   0   'False
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   120
      TabIndex        =   3
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
      Image           =   "AgingNew2.frx":1ED0
      cBack           =   -2147483633
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "AGING PERIOD:"
      Height          =   255
      Left            =   10440
      TabIndex        =   20
      Top             =   150
      Width           =   1695
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   720
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   15075
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuAging 
      Caption         =   "FlxAging"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "AgingNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAcct As New ADODB.Recordset
Dim rsPrn As New ADODB.Recordset


Private Sub cboCType_Change()
    Dim rsAging As New ADODB.Recordset
    Dim Qry As String
    
    If cboCType.Text = "ALL CUSTOMER" Then
        Qry = "SELECT agingcurrent.CUSTOMER, agingcurrent.CREDITLIMIT, agingcurrent.REFNO, agingcurrent.REFDATE, agingcurrent.REFAMT, agingcurrent.TERMS, agingcurrent.DUEDATE, agingcurrent.CURRENT, agingcurrent.DNXTMO, agingcurrent.OVER1MO, agingcurrent.OVER2MO, " & _
              "agingcurrent.OVER3MO, agingcurrent.TOTALDUE, agingcurrent.PAYMENT, agingcurrent.PRETURN, agingcurrent.CM, agingcurrent.TTLCREDIT, agingcurrent.TTLCOLLNXTM, agingcurrent.BRANCH, agingcurrent.AGENT, agingcurrent.CTYPE, agingcurrent.CUSTID, agingcurrent.SPSORDERID " & _
              "From agingcurrent Order By agingcurrent.CUSTOMER, agingcurrent.BRANCH, agingcurrent.REFDATE Desc"
        txtSQL.Text = Qry
    ElseIf cboCType.Text = "WHOLESALE" Then
        Qry = "SELECT agingcurrent.CUSTOMER, agingcurrent.CREDITLIMIT, agingcurrent.REFNO, agingcurrent.REFDATE, agingcurrent.REFAMT, agingcurrent.TERMS, agingcurrent.DUEDATE, agingcurrent.CURRENT, agingcurrent.DNXTMO, agingcurrent.OVER1MO, agingcurrent.OVER2MO, " & _
              "agingcurrent.OVER3MO, agingcurrent.TOTALDUE, agingcurrent.PAYMENT, agingcurrent.PRETURN, agingcurrent.CM, agingcurrent.TTLCREDIT, agingcurrent.TTLCOLLNXTM, agingcurrent.BRANCH, agingcurrent.AGENT, agingcurrent.CTYPE, agingcurrent.CUSTID, agingcurrent.SPSORDERID " & _
              "From agingcurrent Where agingcurrent.CTYPE = 'WHOLESALE' Order By agingcurrent.CUSTOMER, agingcurrent.BRANCH, agingcurrent.REFDATE Desc"
        txtSQL.Text = Qry
    Else
        Qry = "SELECT agingcurrent.CUSTOMER, agingcurrent.CREDITLIMIT, agingcurrent.REFNO, agingcurrent.REFDATE, agingcurrent.REFAMT, agingcurrent.TERMS, agingcurrent.DUEDATE, agingcurrent.CURRENT, agingcurrent.DNXTMO, agingcurrent.OVER1MO, agingcurrent.OVER2MO, " & _
              "agingcurrent.OVER3MO, agingcurrent.TOTALDUE, agingcurrent.PAYMENT, agingcurrent.PRETURN, agingcurrent.CM, agingcurrent.TTLCREDIT, agingcurrent.TTLCOLLNXTM, agingcurrent.BRANCH, agingcurrent.AGENT, agingcurrent.CTYPE, agingcurrent.CUSTID, agingcurrent.SPSORDERID " & _
              "From agingcurrent Where agingcurrent.CTYPE = '" & cboCType.Text & "' Order By agingcurrent.CUSTOMER, agingcurrent.BRANCH, agingcurrent.REFDATE Desc"
        txtSQL.Text = Qry       'Where agingcurrent.CTYPE = '" & cboCType.Text & "'
        
    End If
    
    'FLXAGING:1-CUSTOMER NAME,2-CREDIT LIMIT, 3-REF. NO.,4-REF. DATE, 5-REF. AMT., 6-TERMS, 7-DUE DATE
    '   8-CURRENT, 9-DUE NEXT MO., 10-1MO. OVERDUE, 11-2MOS. OVERDUE, 12-3MOS.& UP, 13-TOTAL OVERDUE,
    '   14-PAYMENTS, 15-RETURNS, 16-CM, 17-TOTAL CREDITS, 18-TOTAL COLL. NXT MO., 19-BRANCH, 20-AGENT, 19-CUST. TYPE, 20-CUST. ID, 21-SPSORDERID
    With rsAging
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Pbar.Value = 0
            Pbar.Max = .RecordCount
            Pbar.Visible = True
            FlxAging.Rows = 1
            Do While Not .EOF
                FlxAging.AddItem vbTab & !CUSTOMER & vbTab & Format(!CreditLimit, "#,##0.00") & vbTab & !RefNo & vbTab & Format(!REFDATE, "MM/DD/YYYY") & vbTab & Format(!RefAmt, "#,##0.00") & _
                vbTab & !Terms & vbTab & Format(!DUEDATE, "MM/DD/YYYY") & vbTab & Format(!Current, "#,##0.00") & vbTab & Format(!DNXTMO, "#,##0.00") & vbTab & Format(!OVER1MO, "#,##0.00") & _
                vbTab & Format(!OVER2MO, "#,##0.00") & vbTab & Format(!OVER3MO, "#,##0.00") & vbTab & Format(!TOTALDUE, "#,##0.00") & vbTab & Format(!PAYMENT, "#,##0.00") & vbTab & Format(!PRETURN, "#,##0.00") & _
                vbTab & Format(!CM, "#,##0.00") & vbTab & Format(!TtlCredit, "#,##0.00") & vbTab & Format(!TTLCOLLNXTM, "#,##0.00") & vbTab & !Branch & vbTab & !Agent & vbTab & !Ctype & vbTab & !CustID & vbTab & !SPSORDERID
                .MoveNext
                Pbar.Value = Pbar.Value + 1
                DoEvents
            Loop
            cmdCReprocess.Visible = True
        Else
            FlxAging.Rows = 1
            FlxAging.Rows = 2
            cmdCReprocess.Visible = True
        End If
    End With
    Pbar.Visible = False
                
End Sub

Private Sub cboCType_Click()
    Call cboCType_Change
End Sub

Private Sub cboCTypSum_Change()
    Dim rsAgeSum As New ADODB.Recordset
    Dim Qry As String
    
    If cboCTypSum.Text = "WHOLESALE" Then
        Qry = "SELECT agingcurrent.CUSTOMER, COUNT(agingcurrent.REFNO) AS TTLINVOICE, agingcurrent.CREDITLIMIT, SUM(agingcurrent.REFAMT) AS SUMREFAMT, SUM(agingcurrent.CURRENT) AS SUMCURRENT, SUM(agingcurrent.DNXTMO) AS SUMDNXTMO, SUM(agingcurrent.OVER1MO) AS SUMOVER1MO, SUM(agingcurrent.OVER2MO) AS SUMOVER2MO, " & _
              "SUM(agingcurrent.OVER3MO) AS SUMOVER3MO, SUM(agingcurrent.TOTALDUE) AS SUMTTLOVERDUE, SUM(agingcurrent.PAYMENT) AS SUMPAYMENTS, SUM(agingcurrent.PRETURN) AS SUMRETURN, SUM(agingcurrent.CM) AS SUMCM, SUM(agingcurrent.TTLCOLLNXTM) AS OUTSTANDINGBAL, agingcurrent.AGENT " & _
              "From agingcurrent WHERE agingcurrent.CTYPE = 'WHOLESALE' Group By agingcurrent.CUSTOMER, agingcurrent.CREDITLIMIT, agingcurrent.AGENT Order By agingcurrent.AGENT, agingcurrent.CUSTOMER"
        
        txtSQLsum.Text = Qry
    Else
        Qry = "SELECT agingcurrent.CUSTOMER, COUNT(agingcurrent.REFNO) AS TTLINVOICE, agingcurrent.CREDITLIMIT, SUM(agingcurrent.REFAMT) AS SUMREFAMT, SUM(agingcurrent.CURRENT) AS SUMCURRENT, SUM(agingcurrent.DNXTMO) AS SUMDNXTMO, SUM(agingcurrent.OVER1MO) AS SUMOVER1MO, SUM(agingcurrent.OVER2MO) AS SUMOVER2MO, " & _
              "SUM(agingcurrent.OVER3MO) AS SUMOVER3MO, SUM(agingcurrent.TOTALDUE) AS SUMTTLOVERDUE, SUM(agingcurrent.PAYMENT) AS SUMPAYMENTS, SUM(agingcurrent.PRETURN) AS SUMRETURN, SUM(agingcurrent.CM) AS SUMCM, SUM(agingcurrent.TTLCOLLNXTM) AS OUTSTANDINGBAL, agingcurrent.AGENT " & _
              "From agingcurrent WHERE agingcurrent.CTYPE = '" & cboCTypSum.Text & "' Group By agingcurrent.CUSTOMER, agingcurrent.CREDITLIMIT, agingcurrent.AGENT Order By agingcurrent.AGENT, agingcurrent.CUSTOMER"
        
        txtSQLsum.Text = Qry
    
    End If
    
    'flxagesum:1-CUSTOMER, 2-TTLINVOICE, 3-CREDITLIMIT, 4-SUMREFAMT, 5-SUMCURRENT, 6-SUMDNXTMO, 7-SUMOVER1MO, 8-SUMOVER2MO, " & _
    '        9-SUMOVER3MO, 10-SUMTTLOVERDUE, 11-SUMPAYMENTS, 12-SUMRETURN, 13-SUMCM, 14-OUTSTANDINGBAL, 15-AGENT "
    With rsAgeSum
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Pbar.Value = 0
            Pbar.Max = .RecordCount
            Pbar.Visible = True
            flxAgeSum.Rows = 1
            Do While Not .EOF
                flxAgeSum.AddItem vbTab & !CUSTOMER & vbTab & Format(!TTLINVOICE, "#,##0") & vbTab & Format(!CreditLimit, "#,##0.00") & vbTab & Format(!SUMRefAMT, "#,##0.00") & _
                vbTab & Format(!SUMCurrent, "#,##0.00") & vbTab & Format(!SUMDNXTMO, "#,##0.00") & vbTab & Format(!SUMOVER1MO, "#,##0.00") & _
                vbTab & Format(!SUMOVER2MO, "#,##0.00") & vbTab & Format(!SUMOVER3MO, "#,##0.00") & vbTab & Format(!SUMTTLOVERDUE, "#,##0.00") & vbTab & Format(!SUMPAYMENTS, "#,##0.00") & vbTab & Format(!SUMRETURN, "#,##0.00") & _
                vbTab & Format(!SUMCM, "#,##0.00") & vbTab & Format(!SUMPAYMENTS + !SUMRETURN + !SUMCM, "#,##0.00") & vbTab & Format(!OUTSTANDINGBAL, "#,##0.00") & vbTab & !Agent
                .MoveNext
                Pbar.Value = Pbar.Value + 1
                DoEvents
            Loop
            cmdPrnSum.Visible = True
        Else
            flxAgeSum.Rows = 1
            flxAgeSum.Rows = 2
            cmdPrnSum.Visible = False
        End If
    End With
    Pbar.Visible = False

End Sub

Private Sub cboCTypSum_Click()
    Call cboCTypSum_Change
End Sub

Private Sub cmdCReprocess_Click()
    Dim rsARdb As New ADODB.Connection
    Dim Qry As String, i As Integer, Ans As Integer
    
    If cboCType.ListIndex = -1 Or cboCType.Text = "" Then
        MsgBox "Please select a Customer Type from the list!", vbInformation + vbOKOnly
        Exit Sub
    End If
    If cboCType.ListIndex = 0 Or cboCType.Text = "ALL CUSTOMER" Then
        Qry = "SELECT customer.CUSTOMERNAME, onaccterms.SPINVNO, onaccterms.TERMS, onaccterms.`STATUS`, onaccterms.BALANCE, onaccterms.ACCTAMOUNT, onaccterms.TRANSDATE, customer.CTYPE, " & _
              "customer.AGENT, customer.CRLIMIT, customer.CUSTOMERID, onaccterms.JOBORDERNO, onaccterms.JOBORDERID, onaccterms.SPSORDERID " & _
              "From customer INNER JOIN onaccterms ON (customer.CUSTOMERID = onaccterms.CUSTOMERID) Where onaccterms.`STATUS` = 'OPEN' Order By customer.CUSTOMERNAME;"
    Else
        Qry = "SELECT customer.CUSTOMERNAME, onaccterms.SPINVNO, onaccterms.TERMS, onaccterms.`STATUS`, onaccterms.BALANCE, onaccterms.ACCTAMOUNT, onaccterms.TRANSDATE, customer.CTYPE, " & _
          "customer.AGENT, customer.CRLIMIT, customer.CUSTOMERID, onaccterms.JOBORDERNO, onaccterms.JOBORDERID, onaccterms.SPSORDERID " & _
          "From customer INNER JOIN onaccterms ON (customer.CUSTOMERID = onaccterms.CUSTOMERID) Where customer.CTYPE = '" & cboCType.Text & "' AND onaccterms.`STATUS` = 'OPEN' Order By customer.CUSTOMERNAME;"
    End If
    'flxacct = 1-CUSTOMERNAME, 2-SPINVNO, 3-TERMS, 4-`STATUS`, 5-BALANCE, 6-ACCTAMOUNT, 7-TRANSDATE, 8-CTYPE,
    '9-AGENT, 10-CRLIMIT, 11-CUSTOMERID, 12-JOBORDERNO, 13-JOBORDERID, 14-SPSORDERID

    If rsAcct.State = adStateOpen Then rsAcct.Close
    rsAcct.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxAccts.DataSource = rsAcct
    flxAccts.Refresh
    If rsAcct.RecordCount = 0 Then
        MsgBox "No Credit found for customer type " & cboCType.Text & "!", vbInformation + vbOKOnly, "Record not found"
        Exit Sub
    End If
    
    Frame2.Enabled = True
    FlxAging.Rows = 1
    Pbar.Value = 0
    Pbar.Max = flxAccts.Rows
    Pbar.Visible = True
    With flxAccts
        For i = 1 To .Rows - 1
            FlxAging.AddItem vbTab & .TextMatrix(i, 1)
            Call ComputeAging(i, FlxAging.Rows - 1)
            FlxAging.TextMatrix(FlxAging.Rows - 1, 21) = cboCType.Text
            If SysBranch = "LYR MOTORPARTS CORP. - WHOLESALE" Then
                FlxAging.TextMatrix(FlxAging.Rows - 1, 19) = "LYR MOTORPARTS CORP. - WHOLESALE"
            ElseIf SysBranch = "LYR MOTORPARTS CORP. - SOBRECARY" Then
                FlxAging.TextMatrix(FlxAging.Rows - 1, 19) = "LYR MOTORPARTS CORP. - SOBRECARY"
            End If
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    Pbar.Visible = True
    If SysBranch = "LYR MOTORPARTS CORP. - WHOLESALE" Then
        If ConnectMRC1 Then 'Process Aging for MRC2
            Qry = "SELECT customer.CUSTOMERNAME, onaccterms.SPINVNO, onaccterms.TERMS, onaccterms.`STATUS`, onaccterms.BALANCE, onaccterms.ACCTAMOUNT, onaccterms.TRANSDATE, customer.CTYPE, " & _
                  "customer.AGENT, customer.CRLIMIT, customer.CUSTOMERID, onaccterms.JOBORDERNO, onaccterms.JOBORDERID, onaccterms.SPSORDERID " & _
                  "From customer INNER JOIN onaccterms ON (customer.CUSTOMERID = onaccterms.CUSTOMERID) Where customer.CTYPE = '" & cboCType.Text & "' AND onaccterms.`STATUS` = 'OPEN' Order By customer.CUSTOMERNAME;"
            
            'flxacct = 1-CUSTOMERNAME, 2-SPINVNO, 3-TERMS, 4-`STATUS`, 5-BALANCE, 6-ACCTAMOUNT, 7-TRANSDATE, 8-CTYPE,
            '9-AGENT, 10-CRLIMIT, 11-CUSTOMERID, 12-JOBORDERNO, 13-JOBORDERID, 14-SPSORDERID

            If rsAcct.State = adStateOpen Then rsAcct.Close
            rsAcct.Open Qry, conMRC1, adOpenKeyset, adLockOptimistic
            Set flxAccts.DataSource = rsAcct
            flxAccts.Refresh
            If rsAcct.RecordCount > 0 Then
                Pbar.Value = 0
                Pbar.Max = flxAccts.Rows
                Pbar.Visible = True
                With flxAccts
                    For i = 1 To .Rows - 1
                        FlxAging.AddItem vbTab & .TextMatrix(i, 1)
                        Call ComputeAgingBranch(i, FlxAging.Rows - 1, "MRC1")
                        FlxAging.TextMatrix(FlxAging.Rows - 1, 19) = "LYR MOTORPARTS CORP. - SOBRECARY"
                        FlxAging.TextMatrix(FlxAging.Rows - 1, 21) = cboCType.Text
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Next i
                End With
                Pbar.Visible = True
            End If
        End If
        If ConnectVAN1 Then 'Process Aging for MRC2
            Qry = "SELECT customer.CUSTOMERNAME, onaccterms.SPINVNO, onaccterms.TERMS, onaccterms.`STATUS`, onaccterms.BALANCE, onaccterms.ACCTAMOUNT, onaccterms.TRANSDATE, customer.CTYPE, " & _
                  "customer.AGENT, customer.CRLIMIT, customer.CUSTOMERID, onaccterms.JOBORDERNO, onaccterms.JOBORDERID, onaccterms.SPSORDERID " & _
                  "From customer INNER JOIN onaccterms ON (customer.CUSTOMERID = onaccterms.CUSTOMERID) Where customer.CTYPE = '" & cboCType.Text & "' AND onaccterms.`STATUS` = 'OPEN' Order By customer.CUSTOMERNAME;"
            
            'flxacct = 1-CUSTOMERNAME, 2-SPINVNO, 3-TERMS, 4-`STATUS`, 5-BALANCE, 6-ACCTAMOUNT, 7-TRANSDATE, 8-CTYPE,
            '9-AGENT, 10-CRLIMIT, 11-CUSTOMERID, 12-JOBORDERNO, 13-JOBORDERID, 14-SPSORDERID

            If rsAcct.State = adStateOpen Then rsAcct.Close
            rsAcct.Open Qry, conVAN1, adOpenKeyset, adLockOptimistic
            Set flxAccts.DataSource = rsAcct
            flxAccts.Refresh
            If rsAcct.RecordCount > 0 Then
                Pbar.Value = 0
                Pbar.Max = flxAccts.Rows
                Pbar.Visible = True
                With flxAccts
                    For i = 1 To .Rows - 1
                        FlxAging.AddItem vbTab & .TextMatrix(i, 1)
                        Call ComputeAgingBranch(i, FlxAging.Rows - 1, "VAN1")
                        FlxAging.TextMatrix(FlxAging.Rows - 1, 19) = "VAN1"
                        FlxAging.TextMatrix(FlxAging.Rows - 1, 21) = cboCType.Text
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Next i
                End With
                Pbar.Visible = True
            End If
        End If
        If ConnectVAN2 Then 'Process Aging for MRC2
            Qry = "SELECT customer.CUSTOMERNAME, onaccterms.SPINVNO, onaccterms.TERMS, onaccterms.`STATUS`, onaccterms.BALANCE, onaccterms.ACCTAMOUNT, onaccterms.TRANSDATE, customer.CTYPE, " & _
                  "customer.AGENT, customer.CRLIMIT, customer.CUSTOMERID, onaccterms.JOBORDERNO, onaccterms.JOBORDERID, onaccterms.SPSORDERID " & _
                  "From customer INNER JOIN onaccterms ON (customer.CUSTOMERID = onaccterms.CUSTOMERID) Where customer.CTYPE = '" & cboCType.Text & "' AND onaccterms.`STATUS` = 'OPEN' Order By customer.CUSTOMERNAME;"
            
            'flxacct = 1-CUSTOMERNAME, 2-SPINVNO, 3-TERMS, 4-`STATUS`, 5-BALANCE, 6-ACCTAMOUNT, 7-TRANSDATE, 8-CTYPE,
            '9-AGENT, 10-CRLIMIT, 11-CUSTOMERID, 12-JOBORDERNO, 13-JOBORDERID, 14-SPSORDERID

            If rsAcct.State = adStateOpen Then rsAcct.Close
            rsAcct.Open Qry, conVAN2, adOpenKeyset, adLockOptimistic
            Set flxAccts.DataSource = rsAcct
            flxAccts.Refresh
            If rsAcct.RecordCount > 0 Then
                Pbar.Value = 0
                Pbar.Max = flxAccts.Rows
                Pbar.Visible = True
                With flxAccts
                    For i = 1 To .Rows - 1
                        FlxAging.AddItem vbTab & .TextMatrix(i, 1)
                        Call ComputeAgingBranch(i, FlxAging.Rows - 1, "VAN2")
                        FlxAging.TextMatrix(FlxAging.Rows - 1, 19) = "VAN2"
                        FlxAging.TextMatrix(FlxAging.Rows - 1, 21) = cboCType.Text
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Next i
                End With
                Pbar.Visible = True
            End If
        End If
'        If ConnectVan3 Then 'Process Aging for MRC2
'            Qry = "SELECT customer.CUSTOMERNAME, onaccterms.SPINVNO, onaccterms.TERMS, onaccterms.`STATUS`, onaccterms.BALANCE, onaccterms.ACCTAMOUNT, onaccterms.TRANSDATE, customer.CTYPE, " & _
'                  "customer.AGENT, customer.CRLIMIT, customer.CUSTOMERID, onaccterms.JOBORDERNO, onaccterms.JOBORDERID, onaccterms.SPSORDERID " & _
'                  "From customer INNER JOIN onaccterms ON (customer.CUSTOMERID = onaccterms.CUSTOMERID) Where customer.CTYPE = '" & cboCType.Text & "' AND onaccterms.`STATUS` = 'OPEN' Order By customer.CUSTOMERNAME;"
'
'            'flxacct = 1-CUSTOMERNAME, 2-SPINVNO, 3-TERMS, 4-`STATUS`, 5-BALANCE, 6-ACCTAMOUNT, 7-TRANSDATE, 8-CTYPE,
'            '9-AGENT, 10-CRLIMIT, 11-CUSTOMERID, 12-JOBORDERNO, 13-JOBORDERID, 14-SPSORDERID
'
'            If rsAcct.State = adStateOpen Then rsAcct.Close
'            rsAcct.Open Qry, conVAN3, adOpenKeyset, adLockOptimistic
'            Set flxAccts.DataSource = rsAcct
'            flxAccts.Refresh
'            If rsAcct.RecordCount > 0 Then
'                Pbar.Value = 0
'                Pbar.Max = flxAccts.Rows
'                Pbar.Visible = True
'                With flxAccts
'                    For i = 1 To .Rows - 1
'                        FlxAging.AddItem vbTab & .TextMatrix(i, 1)
'                        Call ComputeAgingBranch(i, FlxAging.Rows - 1, "VAN3")
'                        FlxAging.TextMatrix(FlxAging.Rows - 1, 19) = "VAN3"
'                        FlxAging.TextMatrix(FlxAging.Rows - 1, 21) = cboCType.Text
'                        Pbar.Value = Pbar.Value + 1
'                        DoEvents
'                    Next i
'                End With
'                Pbar.Visible = True
'            End If
'        End If
    End If
    Pbar.Visible = False
    Qry = "SELECT COUNT(agingcurrent.CUSTOMER) As FIELD_1 From agingcurrent Where agingcurrent.CTYPE = '" & cboCType.Text & "'"
    If Val(GetFieldValue(Qry)) >= 1 Then
        Ans = MsgBox("A previous record of " & cboCType.Text & " AGING already exist!" & vbCrLf & "Do you want to overwrite Aging Records?", vbQuestion + vbYesNo, "Save")
        If Ans = vbNo Then Exit Sub
        LYRcon.Execute "DELETE From agingcurrent Where agingcurrent.CTYPE = '" & cboCType.Text & "';"
    Else
        Ans = MsgBox("Do you want to save Aging Records?", vbQuestion + vbYesNo, "Save")
    End If
    If Ans = vbNo Then Exit Sub
    With FlxAging
        Pbar.Value = 0
        Pbar.Max = .Rows
        Pbar.Visible = True
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 3) <> "" Then
                LYRcon.Execute "INSERT INTO agingcurrent(CUSTOMER, BRANCH, AGENT, CTYPE, CREDITLIMIT, CUSTID, REFNO, REFDATE, REFAMT, TERMS, DUEDATE, CURRENT, DNXTMO, OVER1MO, OVER2MO, OVER3MO, TOTALDUE, PAYMENT, PRETURN, CM, TTLCREDIT, TTLCOLLNXTM, SPSORDERID) " & _
                "VALUES('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 19) & "', '" & .TextMatrix(i, 20) & "', '" & .TextMatrix(i, 21) & "', " & Val(.TextMatrix(i, 2)) & ", " & Val(.TextMatrix(i, 22)) & ", '" & .TextMatrix(i, 3) & "', " & _
                "'" & Format(DateValue(.TextMatrix(i, 4)), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 5)) & ", " & Val(.TextMatrix(i, 6)) & ", '" & Format(DateValue(.TextMatrix(i, 7)), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 8)) & ", " & Val(.TextMatrix(i, 9)) & _
                ", " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 11)) & ", " & Val(.TextMatrix(i, 12)) & ", " & Val(.TextMatrix(i, 13)) & ", " & Val(.TextMatrix(i, 14)) & ", " & Val(.TextMatrix(i, 15)) & ", " & Val(.TextMatrix(i, 16)) & ", " & Val(.TextMatrix(i, 17)) & ", " & Val(.TextMatrix(i, 18)) & ", " & Val(.TextMatrix(i, 19)) & ");"
                End If
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    Pbar.Visible = False
    MsgBox "Aging successfully saved"
    
        
            
End Sub

Sub ComputeAging(AcctRow As Integer, AgingRow As Integer)
    Dim Qry As String
    Dim Cr As Currency, Dr As Currency, Balance As Currency
    
    'Identify transaction if Job Order
    If Val(flxAccts.TextMatrix(AcctRow, 13)) > 0 And flxAccts.TextMatrix(AcctRow, 2) = "" Then 'JobOrderID>0 and InvNo=""
        'Process Job Order
        'Get Payment
        Qry = "SELECT SUM(Ledger.CREDIT) As TtlPay From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 12) & "' AND ledger.REFTYPE = 'Payment' AND ledger.JOBID = " & Val(flxAccts.TextMatrix(AcctRow, 13)) & ";"
        FlxAging.TextMatrix(AgingRow, 14) = Format(Val(GetFieldValue(Qry)), "###0.00")
        'Get Credit from DMCM
        Qry = "SELECT SUM(Ledger.CREDIT) As TtlCr From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 12) & "' AND ledger.REFTYPE = 'DMCM' AND ledger.JOBID = " & Val(flxAccts.TextMatrix(AcctRow, 13)) & ";"
        Cr = Val(GetFieldValue(Qry))
        Qry = "SELECT SUM(Ledger.DEBIT) As TtlCr From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 12) & "' AND ledger.REFTYPE = 'DMCM' AND ledger.JOBID = " & Val(flxAccts.TextMatrix(AcctRow, 13)) & ";"
        Dr = Val(GetFieldValue(Qry))
        If Dr > 0 Then Dr = Dr * -1
        FlxAging.TextMatrix(AgingRow, 16) = Format(Cr + Dr, "###0.00")  'Display DMCM
        FlxAging.TextMatrix(AgingRow, 17) = Val(FlxAging.TextMatrix(AgingRow, 14)) + Val(FlxAging.TextMatrix(AgingRow, 15)) + Val(FlxAging.TextMatrix(AgingRow, 16))
        Balance = Val(flxAccts.TextMatrix(AcctRow, 6)) - Val(FlxAging.TextMatrix(AgingRow, 17))
        FlxAging.TextMatrix(AgingRow, 2) = Format(Val(flxAccts.TextMatrix(AcctRow, 10)), "###0.00") 'Credit Limit
        FlxAging.TextMatrix(AgingRow, 3) = "JO# " & flxAccts.TextMatrix(AcctRow, 12)    'Ref No
        FlxAging.TextMatrix(AgingRow, 4) = Format(flxAccts.TextMatrix(AcctRow, 7), "mm/dd/yyyy") 'Ref Date
        FlxAging.TextMatrix(AgingRow, 5) = Format(Val(flxAccts.TextMatrix(AcctRow, 6)), "###0.00") 'Ref Amt
        FlxAging.TextMatrix(AgingRow, 6) = flxAccts.TextMatrix(AcctRow, 3)    'Term
        FlxAging.TextMatrix(AgingRow, 7) = GetDueDate(DateValue(FlxAging.TextMatrix(AgingRow, 4)), Val(FlxAging.TextMatrix(AgingRow, 6)))   'Due Date
        If Date > DateValue(FlxAging.TextMatrix(AgingRow, 4)) And Date < DateValue(FlxAging.TextMatrix(AgingRow, 7)) Then 'date within invdate to duedate
            FlxAging.TextMatrix(AgingRow, 8) = Format(Balance, "###0.00") 'Current
        ElseIf Date >= DateValue(FlxAging.TextMatrix(AgingRow, 7)) And Date < DateAdd("M", 1, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate and <=duedate + 1 Month
            FlxAging.TextMatrix(AgingRow, 9) = Format(Balance, "###0.00") 'Due Next Month
        ElseIf Date >= DateAdd("M", 1, DateValue(FlxAging.TextMatrix(AgingRow, 7))) And Date < DateAdd("M", 2, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate+1Mo and <=duedate + 2 Months
            FlxAging.TextMatrix(AgingRow, 10) = Format(Balance, "###0.00") '1 Month Overdue
        ElseIf Date >= DateAdd("M", 2, DateValue(FlxAging.TextMatrix(AgingRow, 7))) And Date < DateAdd("M", 3, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate+2Mo and <=duedate + 3 Months
            FlxAging.TextMatrix(AgingRow, 11) = Format(Balance, "###0.00") '2 Months Overdue
        ElseIf Date >= DateAdd("M", 3, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate + 3 Mos
            FlxAging.TextMatrix(AgingRow, 12) = Format(Balance, "###0.00") '3 Months Up Overdue
        End If
        FlxAging.TextMatrix(AgingRow, 13) = Format(Val(FlxAging.TextMatrix(AgingRow, 10)) + Val(FlxAging.TextMatrix(AgingRow, 11)) + Val(FlxAging.TextMatrix(AgingRow, 12)), "###0.00") 'Total Overdue
        FlxAging.TextMatrix(AgingRow, 18) = Format(Balance, "###0.00") 'Total Coll Nxt Mo.
        FlxAging.TextMatrix(AgingRow, 20) = flxAccts.TextMatrix(AcctRow, 9) 'Agent
        FlxAging.TextMatrix(AgingRow, 22) = flxAccts.TextMatrix(AcctRow, 11) 'Customer ID
        FlxAging.TextMatrix(AgingRow, 23) = flxAccts.TextMatrix(AcctRow, 13) 'jOBORder ID
        
    
    Else    'Process Invoice
        'Get Payment
        Qry = "SELECT SUM(Ledger.CREDIT) As TtlPay From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 2) & "' AND ledger.REFTYPE = 'Payment'"
        FlxAging.TextMatrix(AgingRow, 14) = Format(Val(GetFieldValue(Qry)), "###0.00")
        'Get Return
        Qry = "SELECT SUM(Ledger.CREDIT) As TtlReturn From Ledger Where ledger.REFTYPE = 'Sales Return' AND ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 2) & "';"
        FlxAging.TextMatrix(AgingRow, 15) = Format(Val(GetFieldValue(Qry)), "###0.00")
        'Get Credit from DMCM
        Qry = "SELECT SUM(Ledger.CREDIT) As TtlCr From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 2) & "' AND ledger.REFTYPE = 'DMCM'"
        Cr = Val(GetFieldValue(Qry))
        Qry = "SELECT SUM(Ledger.DEBIT) As TtlCr From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 2) & "' AND ledger.REFTYPE = 'DMCM'"
        Dr = Val(GetFieldValue(Qry))
        If Dr > 0 Then Dr = Dr * -1
        FlxAging.TextMatrix(AgingRow, 16) = Format(Cr + Dr, "###0.00")  'Display DMCM
        FlxAging.TextMatrix(AgingRow, 17) = Val(FlxAging.TextMatrix(AgingRow, 14)) + Val(FlxAging.TextMatrix(AgingRow, 15)) + Val(FlxAging.TextMatrix(AgingRow, 16))
        Balance = Val(flxAccts.TextMatrix(AcctRow, 6)) - Val(FlxAging.TextMatrix(AgingRow, 17))
        FlxAging.TextMatrix(AgingRow, 2) = Format(Val(flxAccts.TextMatrix(AcctRow, 10)), "###0.00") 'Credit Limit
        FlxAging.TextMatrix(AgingRow, 3) = "Inv# " & flxAccts.TextMatrix(AcctRow, 2)    'Ref No
        FlxAging.TextMatrix(AgingRow, 4) = Format(flxAccts.TextMatrix(AcctRow, 7), "mm/dd/yyyy") 'Ref Date
        FlxAging.TextMatrix(AgingRow, 5) = Format(Val(flxAccts.TextMatrix(AcctRow, 6)), "###0.00") 'Ref Amt
        FlxAging.TextMatrix(AgingRow, 6) = flxAccts.TextMatrix(AcctRow, 3)    'Term
        FlxAging.TextMatrix(AgingRow, 7) = GetDueDate(DateValue(FlxAging.TextMatrix(AgingRow, 4)), Val(FlxAging.TextMatrix(AgingRow, 6)))   'Due Date
        If Date > DateValue(FlxAging.TextMatrix(AgingRow, 4)) And Date < DateValue(FlxAging.TextMatrix(AgingRow, 7)) Then 'date within invdate to duedate
            FlxAging.TextMatrix(AgingRow, 8) = Format(Balance, "###0.00") 'Current
        ElseIf Date >= DateValue(FlxAging.TextMatrix(AgingRow, 7)) And Date < DateAdd("M", 1, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate and <=duedate + 1 Month
            FlxAging.TextMatrix(AgingRow, 9) = Format(Balance, "###0.00") 'Due Next Month
        ElseIf Date >= DateAdd("M", 1, DateValue(FlxAging.TextMatrix(AgingRow, 7))) And Date < DateAdd("M", 2, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate+1Mo and <=duedate + 2 Months
            FlxAging.TextMatrix(AgingRow, 10) = Format(Balance, "###0.00") '1 Month Overdue
        ElseIf Date >= DateAdd("M", 2, DateValue(FlxAging.TextMatrix(AgingRow, 7))) And Date < DateAdd("M", 3, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate+2Mo and <=duedate + 3 Months
            FlxAging.TextMatrix(AgingRow, 11) = Format(Balance, "###0.00") '2 Months Overdue
        ElseIf Date >= DateAdd("M", 3, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate + 3 Mos
            FlxAging.TextMatrix(AgingRow, 12) = Format(Balance, "###0.00") '3 Months Up Overdue
        End If
        FlxAging.TextMatrix(AgingRow, 13) = Format(Val(FlxAging.TextMatrix(AgingRow, 10)) + Val(FlxAging.TextMatrix(AgingRow, 11)) + Val(FlxAging.TextMatrix(AgingRow, 12)), "###0.00") 'Total Overdue
        FlxAging.TextMatrix(AgingRow, 18) = Format(Balance, "###0.00") 'Total Coll Nxt Mo.
        FlxAging.TextMatrix(AgingRow, 20) = flxAccts.TextMatrix(AcctRow, 9) 'Agent
        FlxAging.TextMatrix(AgingRow, 22) = flxAccts.TextMatrix(AcctRow, 11) 'Customer ID
        FlxAging.TextMatrix(AgingRow, 23) = flxAccts.TextMatrix(AcctRow, 14) 'SPsalesORder ID
        
    End If
End Sub

Sub ComputeAgingBranch(AcctRow As Integer, AgingRow As Integer, myCON As String)
    Dim Qry As String
    Dim Cr As Currency, Dr As Currency, Balance As Currency
    
    'Identify transaction if Job Order
    If Val(flxAccts.TextMatrix(AcctRow, 13)) > 0 And flxAccts.TextMatrix(AcctRow, 2) = "" Then 'JobOrderID>0 and InvNo=""
        'Process Job Order
        
    Else    'Process Invoice
        'Get Payment
        Qry = "SELECT SUM(Ledger.CREDIT) As TtlPay From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 2) & "' AND ledger.REFTYPE = 'Payment'"
        FlxAging.TextMatrix(AgingRow, 14) = Format(Val(GetFieldValueBranch(Qry, myCON)), "###0.00")
        'Get Return
        Qry = "SELECT SUM(Ledger.CREDIT) As TtlReturn From Ledger Where ledger.REFTYPE = 'Sales Return' AND ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 2) & "';"
        FlxAging.TextMatrix(AgingRow, 15) = Format(Val(GetFieldValueBranch(Qry, myCON)), "###0.00")
        'Get Credit from DMCM
        Qry = "SELECT SUM(Ledger.CREDIT) As TtlCr From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 2) & "' AND ledger.REFTYPE = 'DMCM'"
        Cr = Val(GetFieldValueBranch(Qry, myCON))
        Qry = "SELECT SUM(Ledger.DEBIT) As TtlCr From Ledger Where ledger.SPINVNO = '" & flxAccts.TextMatrix(AcctRow, 2) & "' AND ledger.REFTYPE = 'DMCM'"
        Dr = Val(GetFieldValueBranch(Qry, myCON))
        If Dr > 0 Then Dr = Dr * -1
        FlxAging.TextMatrix(AgingRow, 16) = Format(Cr + Dr, "###0.00")  'Display DMCM
        FlxAging.TextMatrix(AgingRow, 17) = Val(FlxAging.TextMatrix(AgingRow, 14)) + Val(FlxAging.TextMatrix(AgingRow, 15)) + Val(FlxAging.TextMatrix(AgingRow, 16))
        Balance = Val(flxAccts.TextMatrix(AcctRow, 6)) - Val(FlxAging.TextMatrix(AgingRow, 17))
        FlxAging.TextMatrix(AgingRow, 2) = Format(Val(flxAccts.TextMatrix(AcctRow, 10)), "###0.00") 'Credit Limit
        FlxAging.TextMatrix(AgingRow, 3) = "Inv# " & flxAccts.TextMatrix(AcctRow, 2)    'Ref No
        FlxAging.TextMatrix(AgingRow, 4) = Format(flxAccts.TextMatrix(AcctRow, 7), "mm/dd/yyyy") 'Ref Date
        FlxAging.TextMatrix(AgingRow, 5) = Format(Val(flxAccts.TextMatrix(AcctRow, 6)), "###0.00") 'Ref Amt
        FlxAging.TextMatrix(AgingRow, 6) = flxAccts.TextMatrix(AcctRow, 3)    'Term
        FlxAging.TextMatrix(AgingRow, 7) = GetDueDate(DateValue(FlxAging.TextMatrix(AgingRow, 4)), Val(FlxAging.TextMatrix(AgingRow, 6)))   'Due Date
        If Date > DateValue(FlxAging.TextMatrix(AgingRow, 4)) And Date < DateValue(FlxAging.TextMatrix(AgingRow, 7)) Then 'date within invdate to duedate
            FlxAging.TextMatrix(AgingRow, 8) = Format(Balance, "###0.00") 'Current
        ElseIf Date >= DateValue(FlxAging.TextMatrix(AgingRow, 7)) And Date < DateAdd("M", 1, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate and <=duedate + 1 Month
            FlxAging.TextMatrix(AgingRow, 9) = Format(Balance, "###0.00") 'Due Next Month
        ElseIf Date >= DateAdd("M", 1, DateValue(FlxAging.TextMatrix(AgingRow, 7))) And Date < DateAdd("M", 2, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate+1Mo and <=duedate + 2 Months
            FlxAging.TextMatrix(AgingRow, 10) = Format(Balance, "###0.00") '1 Month Overdue
        ElseIf Date >= DateAdd("M", 2, DateValue(FlxAging.TextMatrix(AgingRow, 7))) And Date < DateAdd("M", 3, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate+2Mo and <=duedate + 3 Months
            FlxAging.TextMatrix(AgingRow, 11) = Format(Balance, "###0.00") '2 Months Overdue
        ElseIf Date >= DateAdd("M", 3, DateValue(FlxAging.TextMatrix(AgingRow, 7))) Then 'date >= duedate + 3 Mos
            FlxAging.TextMatrix(AgingRow, 12) = Format(Balance, "###0.00") '3 Months Up Overdue
        End If
        FlxAging.TextMatrix(AgingRow, 13) = Format(Val(FlxAging.TextMatrix(AgingRow, 10)) + Val(FlxAging.TextMatrix(AgingRow, 11)) + Val(FlxAging.TextMatrix(AgingRow, 12)), "###0.00") 'Total Overdue
        FlxAging.TextMatrix(AgingRow, 18) = Format(Balance, "###0.00") 'Total Coll Nxt Mo.
        FlxAging.TextMatrix(AgingRow, 20) = flxAccts.TextMatrix(AcctRow, 9) 'Agent
        FlxAging.TextMatrix(AgingRow, 22) = flxAccts.TextMatrix(AcctRow, 11) 'Customer ID
        FlxAging.TextMatrix(AgingRow, 23) = flxAccts.TextMatrix(AcctRow, 14) 'SPsalesORder ID
        
    End If
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    'MotoRep.tbMain.Visible = True
    'MotoRep.imgBG.Visible = True
    MotoRep.toolBarButton.Visible = True
End Sub

Private Sub Form_Load()
    Dim rsCType As New ADODB.Recordset
    Dim Qry As String
    
    tabAging.Tab = 0
    'tabAging.TabEnabled(1) = False
    Qry = "SELECT DISTINCT customer.CTYPE From customer Order By customer.CTYPE Desc;"
    With rsCType
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboCType.AddItem "ALL CUSTOMER"
            Do While Not .EOF
                cboCType.AddItem !Ctype
                cboCTypSum.AddItem !Ctype
                .MoveNext
            Loop
            cboCType.ListIndex = 0
            cboCTypSum.ListIndex = 0
        End If
        .Close
    End With
    
    cboAgingPeriod.AddItem "CURRENT AGING"
    
End Sub

Private Sub tabAging_Click(PreviousTab As Integer)
    Dim crNwAging As New rptAgingNew
    
    
    If PreviousTab = 1 Then Exit Sub
        
    If txtSQL.Text = "" Then
        MsgBox "No Record to display!", vbExclamation + vbOKOnly, "Print Preview"
        Exit Sub
    End If
    
    'Display Print Preview
    If rsPrn.State = adStateOpen Then rsPrn.Close
    rsPrn.Open txtSQL.Text, LYRcon, adOpenDynamic, adLockOptimistic
    crNwAging.DiscardSavedData
    crNwAging.Database.SetDataSource rsPrn
    crNwAging.ParameterFields(1).ClearCurrentValueAndRange
    crNwAging.ParameterFields(1).AddCurrentValue UCase$(SysBranch)
    crNwAging.ParameterFields(2).ClearCurrentValueAndRange
    crNwAging.ParameterFields(2).AddCurrentValue UCase$(SysAddress)
    crNwAging.ParameterFields(3).ClearCurrentValueAndRange
    crNwAging.ParameterFields(3).AddCurrentValue "As of " & Format(Date, "Mmmm dd, yyyy")
    
    CRViewer1.ReportSource = crNwAging
    CRViewer1.ViewReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.EnableExportButton = True
        
End Sub


