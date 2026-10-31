VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MonthlyColl 
   Caption         =   "Monthly Collection Report"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Collection for the Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   5280
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   720
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid flxSDay 
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   280
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
         FormatString    =   "^CASH |^ON-ACCOUNT  |^SERVICE   |^DP  |^OTHERS  |^TOTAL               "
      End
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   8520
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   9150
      Visible         =   0   'False
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   0
   End
   Begin VB.TextBox txtYr 
      Height          =   285
      Left            =   7800
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   6960
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   7335
      Left            =   120
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   1800
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   12938
      _Version        =   393216
      RowHeightMin    =   285
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   5055
      Begin VB.ComboBox cboCutOff 
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   172
         Width           =   3135
      End
      Begin VB.Label lblMoColl 
         Alignment       =   2  'Center
         Caption         =   "Total Collection for the Month"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "For the Month of:"
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
         TabIndex        =   7
         Top             =   210
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "MonthlyColl.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   6
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
      Image           =   "MonthlyColl.frx":059A
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPro 
      Height          =   1095
      Left            =   960
      TabIndex        =   13
      Top             =   9720
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      FormatString    =   "DATE |>CASH|>ONACCT|>SERVICE|>DP|>OTHERS|>TOTAL"
   End
   Begin VB.Label lblTotal 
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
      Height          =   255
      Left            =   13680
      TabIndex        =   14
      Tag             =   "XY"
      Top             =   9150
      Width           =   615
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14670
      m_MDI_Height    =   9720
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHLY COLLECTION REPORT"
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
      Left            =   9720
      TabIndex        =   2
      Top             =   450
      Width           =   4500
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
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   2615
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
      Left            =   150
      TabIndex        =   1
      Top             =   475
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHLY COLLECTION REPORT"
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
      Left            =   9735
      TabIndex        =   3
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuINV 
      Caption         =   "SP Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuProcess 
         Caption         =   "Verify Sales"
      End
   End
End
Attribute VB_Name = "MonthlyColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMSales As New ADODB.Recordset
Dim ActRow As Integer


Private Sub cboCutOff_Click()
    Dim MoYr() As String
    
    If cboCutOff.ListIndex = -1 Or cboCutOff.Text = "" Then Exit Sub
    MoYr = Split(cboCutOff.Text, "-")
    txtYr.Text = Trim(MoYr(1))
    txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & " 1, " & txtYr.Text))
    Call DisplaySales
End Sub


Private Sub cmdTrans_Click(Index As Integer)
    Dim yrTtl As Currency, Qry As String
    Dim dStart As Date, dEnd As Date
    
    Select Case Index
        Case 2  'Print
            dStart = DateSerial(txtYr.Text, 1, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 1, DateSerial(txtYr.Text, txtMo.Text, 1)))
            'Qry = "SELECT Sum(MSALES.TTLCOLL) AS SumOfTTLCOLL From MSALES WHERE (((MSALES.SDATE)>=#" & dStart & "# And (MSALES.SDATE)<=#" & dEnd & "#));"
            'mySQl
            Qry = "call sp_mrcretails.sp_dailyCashReport_perMonthRange('" & sysBranchCode & "','" & Format(dStart, "yyyy-mm-dd") & "','" & Format(dEnd, "yyyy-mm-dd") & "')"
    
            yrTtl = Val(GetFieldValue(Qry))
            PRNmc.MySource = "MONTHLYCOLLECTION"
            PRNmc.QryPrn = yrTtl
            PRNmc.ParamValue1 = "For the Month of " & cboCutOff.Text
            PRNmc.ParamValue2 = "Total Collection from " & Format(dStart, "Mmmm dd, yyyy") & " to " & Format(dEnd, "Mmmm dd, yyyy") & ": " & Format(yrTtl, "#,##0.00")
            PRNmc.ParamValue3 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='CASHIER'));")
            PRNmc.ParamValue4 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='OIC'));")
            PRNmc.ParamValue5 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='MANAGER'));")
            PRNmc.Show
        Case 3  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
    End Select
End Sub


Private Sub flxSP_EnterCell()
    ActRow = flxSP.Row
End Sub


Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuINV)
    End If
End Sub

Private Sub LoadTodaySales()
    Dim rsTSales As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT MSALES.* From MSALES WHERE (((MSALES.SDATE)=#" & Date & "#));"
    'mySQL
    Qry = "SELECT MSALES.* From MSALES WHERE MSALES.SDATE='" & Format(Date, "yyyy-mm-dd") & "';"
    
    With rsTSales
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSDay.TextMatrix(1, 0) = Format(!Cash, "#,##0.00")
            flxSDay.TextMatrix(1, 1) = Format(!ACCTCOLL, "#,##0.00")
            flxSDay.TextMatrix(1, 2) = Format(!SrvcColl, "#,##0.00")
            flxSDay.TextMatrix(1, 3) = Format(!DP, "#,##0.00")
            flxSDay.TextMatrix(1, 4) = Format(!SalesRet, "#,##0.00")
            flxSDay.TextMatrix(1, 5) = Format(!TTLCOLL, "#,##0.00")
        End If
        .Close
    End With
End Sub

Private Sub Form_Load()
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String, Cash As Single, Acct As Single, Srvc As Single, DP As Single, Others As Single, Ttl As Single
    Dim Mo As Integer, Yr As Integer, MoYr As Boolean, TmpD As Date
    'Display Month and year from start of business upto current month
    
    Mo = 1
    Yr = 2008
    Do
        cboCutOff.AddItem MonthName(Mo) & "-" & Yr, 0
        TmpD = DateAdd("m", 1, DateSerial(Yr, Mo, 1))
        Mo = Month(TmpD)
        Yr = Year(TmpD)
        MoYr = IIf(TmpD <= DateSerial(Year(Date), Month(Date), 1), True, False)
    Loop While MoYr
        
    'Qry = "SELECT MSALES.TTLCOLL From MSALES WHERE (((MSALES.TTLCOLL)>0) AND ((MSALES.SDATE)=#" & Date & "#));"
    'mySQL
    Qry = "SELECT MSALES.TTLCOLL From MSALES WHERE MSALES.TTLCOLL>0 AND MSALES.SDATE='" & Format(Date, "yyyy-mm-dd") & "';"
    
    If CtrlNumExist(Qry) Then
        Call LoadTodaySales
    Else
        'Cash
        'Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                 "Where (((PAYMENTS.PAYDATE) = #" & Date & "#) And (Not (PAYMENTS.INVOICENO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
        'mySQL
        Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP From payments " & _
              "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL"
        Cash = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(1, 0) = Format(Cash, "#,##0.00")
        
        'On Account
        'Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                   "Where (((PAYMENTS.PAYDATE) = #" & Date & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
        'mySQL
        Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP From payments " & _
              "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.ORNO IS NOT NULL"
        Acct = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(1, 1) = Format(Acct, "#,##0.00")
        
        'Service
        'Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
                  "Where (((PAYMENTS.PAYDATE) = #" & Date & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SERVICE))<>0));"
        'mySQL
        'Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
                  "Where PAYMENTS.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' And Not (PAYMENTS.ORNO) Is Null HAVING Sum(PAYMENTS.SERVICE)<>0;"
        Qry = "SELECT SUM(payments.SERVICE) As SumOfSVC From payments " & _
              "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.ORNO IS NOT NULL"
        
        
        Srvc = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(1, 2) = Format(Srvc, "#,##0.00")
        'Down Payment
        'Qry = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT From PAYMENTS " & _
                "Where (((PAYMENTS.PAYDATE) = #" & Date & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.DOWNPAYMENT))<>0));"
        'mySQl
        'Qry = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT From PAYMENTS " & _
                "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.DOWNPAYMENT))<>0)));"
        Qry = "SELECT SUM(payments.DOWNPAYMENT) As SumOfDP From payments " & _
              "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.ORNO IS NOT NULL"
        DP = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(1, 3) = Format(DP, "#,##0.00")
        'Others
        'Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
                   "Where (((PAYMENTS.PAYDATE) = #" & Date & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
        'mySQl
        'Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
                   "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
        Qry = "SELECT SUM(payments.Others) As SumOfOTHERS From payments " & _
              "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.ORNO IS NOT NULL"
        Others = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(1, 4) = Format(Others, "#,##0.00")
        Ttl = Cash + Acct + Srvc + DP + Others
        flxSDay.TextMatrix(1, 5) = Format(Ttl, "#,##0.00")
    End If
    Frame2.Caption = "Collection for the day " & Format(Date, "Dddd, Mmmm dd, yyyy")
    
    Timer1.Enabled = True
    
    MotoRep.sBar.Panels(6).Text = "Monthly Collection Report"
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsMSales.State = adStateOpen Then rsMSales.Close
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, j As Integer
    
    lblHeading(2).Left = Frame1.Left + Frame1.Width + (Frame2.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + Frame1.Width + (Frame2.Width - lblHeading(3).Width)
    
    With flxSP             '        1       2        3          4          5         6    7
        .FormatString = "  |^ DATE  |>CASH|>ON-ACCOUNT|>SERVICE|>DOWN PAYMENT |>OTHERS|>TOTAL"
        .ColWidth(0) = 300
        For i = 1 To 7
            .ColWidth(i) = (.Width - 400) / 7
        Next i
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format(.TextMatrix(i, 1), "mm/dd/yyyy")
            For j = 2 To 7
                .TextMatrix(i, j) = Format(.TextMatrix(i, j), "#,##0.00")
            Next j
        Next i
    End With
    With flxSDay
        For i = 0 To 5
            .ColWidth(i) = .Width / 6
        Next i
    End With
End Sub

Private Sub Timer1_Timer()
    cboCutOff.ListIndex = 0
    Timer1.Enabled = False
End Sub


Private Sub DisplaySales()
    Dim Qry As String
    Dim dStart As Date, dEnd As Date
    Dim TtlSP As Currency, TtlSrv As Currency, GTtl As Currency
    
    dStart = DateSerial(txtYr.Text, txtMo.Text, 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
'    'Qry = "SELECT MSALES.SDATE, MSALES.CASH, MSALES.ACCTCOLL, MSALES.SRVCCOLL, MSALES.DP, MSALES.OTHERS, MSALES.TTLCOLL From MSALES " & _
'          "Where (((MSALES.SDATE) >= #" & dStart & "# And (MSALES.SDATE) <= #" & dEnd & "#) And ((MSALES.TTLCOLL) > 0)) ORDER BY MSALES.SDATE;"
'    'mySQl
'    Qry = "SELECT MSALES.SDATE, MSALES.CASH, MSALES.ACCTCOLL, MSALES.SRVCCOLL, MSALES.DP, MSALES.OTHERS, MSALES.TTLCOLL From MSALES " & _
'          "Where (((MSALES.SDATE) >= '" & Format(dStart, "yyyy-mm-dd") & "' And (MSALES.SDATE) <= '" & Format(dEnd, "yyyy-mm-dd") & "') And ((MSALES.TTLCOLL) > 0)) ORDER BY MSALES.SDATE;"
'
'    txtSQL.Text = Qry
'
'    If rsMSales.State = adStateOpen Then rsMSales.Close
'    rsMSales.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'    If rsMSales.RecordCount = 0 Then
'        flxSP.Clear
'        flxSP.Rows = 2
'        flxSP.Cols = 9
'        Call RXResizer1_ResizeComplete
'        Call ProcessSales
'        Exit Sub
'    End If
'    Set flxSP.DataSource = rsMSales
'    ActRow = 1
'    Call RXResizer1_ResizeComplete
'    'Qry = "SELECT Sum(MSALES.TTLCOLL) AS SumOfTTLCOLL From MSALES " & _
'          "WHERE (((MSALES.SDATE)>=#" & dStart & "# And (MSALES.SDATE)<=#" & dEnd & "#));"
'    'mySQL
'    Qry = "SELECT Sum(MSALES.TTLCOLL) AS SumOfTTLCOLL From MSALES " & _
'          "WHERE (((MSALES.SDATE)>='" & Format(dStart, "yyyy-mm-dd") & "' And (MSALES.SDATE)<='" & Format(dEnd, "yyyy-mm-dd") & "'));"
'
'
    
    Qry = "call sp_mrcretails.sp_dailyCashReport_perMonthRange('" & sysBranchCode & "','" & Format(dStart, "yyyy-mm-dd") & "','" & Format(dEnd, "yyyy-mm-dd") & "')"
    
    With rsMSales
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then
            flxSP.Clear
            flxSP.Rows = 2
            flxSP.Cols = 9
            Call RXResizer1_ResizeComplete
            Exit Sub
        End If
        
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 2
            GTtl = 0
            flxSP.Rows = .RecordCount
            For i = 1 To flxSP.Rows - 1
                'PAYDATE, SPCASH, SPCREDIT, SPSERVICE, SPDOWNPAYMENT, SPOTHERS, SPTOTAL
                flxSP.TextMatrix(i, 1) = Format(!Date, "mm/dd/yyyy")
                flxSP.TextMatrix(i, 2) = Format(!Cash, "#,###,##0.00")
                flxSP.TextMatrix(i, 3) = Format(!ACCTCOLL, "#,###,##0.00")
                flxSP.TextMatrix(i, 4) = Format(!Service, "#,###,##0.00")
                flxSP.TextMatrix(i, 5) = Format(!DP, "#,###,##0.00")
                flxSP.TextMatrix(i, 6) = Format(!Others, "#,###,##0.00")
                flxSP.TextMatrix(i, 7) = Format(!SPTOTAL, "#,###,##0.00")
                GTtl = GTtl + Format(!SPTOTAL, "#,###,##0.00")
                'flxSP.AddItem vbTab & Format(!SDATE, "mm/dd/yyyy") & vbTab & Format(!Cash, "#,##0.00") & vbTab & Format(!AcctSales, "#,##0.00") & vbTab & Format(!SalesRet, "#,##0.00") & vbTab & Format(TtlSP, "#,##0.00") & vbTab & Format(!SRVCCASH, "#,##0.00") & vbTab & Format(!SrvcAcct, "#,##0.00") & vbTab & Format(TtlSrv, "#,##0.00") & vbTab & Format(!Others, "#,##0.00") & vbTab & Format(GTtl, "#,##0.00")
                .MoveNext
            Next i
            .Close
        End If
    End With
    ActRow = 1
    Call RXResizer1_ResizeComplete
    
    lblMoColl.Caption = "Total Cash Collection for the Month: " & Format(GTtl, "#,###,##0.00")
    'lblMoColl.Caption = "Total Cash Collection for the Month: " & Format(Val(GetFieldValue(Qry)), "#,##0.00")
End Sub

Private Sub ProcessSales()
    Dim Qry As String, i As Integer, Ans As Integer, pPct As Single
    Dim dStart As Date, dEnd As Date, TmpD As Date
    
    Ans = MsgBox("Do you want to Process Sales for then Period " & cboCutOff.Text & "?", vbQuestion + vbYesNo, "Sales")
    If Ans = vbNo Then Exit Sub
    
    dStart = DateSerial(txtYr.Text, txtMo.Text, 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    
    flxPro.Rows = 1
    TmpD = dStart
    Do While TmpD <= dEnd
        flxPro.AddItem TmpD
        TmpD = DateAdd("d", 1, TmpD)
    Loop
    
    'Lock objects
    Frame1.Enabled = False
    cmdTrans(2).Enabled = False
    cmdTrans(3).Enabled = False
    flxSP.Enabled = False
    
    Pbar.Value = 0
    Pbar.Max = (flxPro.Rows - 1) * 2
    Pbar.Visible = True
    With flxPro
        For i = 1 To .Rows - 1
            'Cash
            'Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                     "Where (((PAYMENTS.PAYDATE) = #" & .TextMatrix(i, 0) & "#) And (Not (PAYMENTS.INVOICENO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
            'mySQl old change these code
            'Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                     "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' And (Not (PAYMENTS.INVOICENO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
            'mySQL
            Qry = "SELECT " & _
                    "Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS  " & _
                "From " & _
                    "PAYMENTS " & _
                "Where PAYMENTS.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' And " & _
                    "PAYMENTS.INVOICENO IS NOT Null HAVING Sum(PAYMENTS.SPAREPARTS)<>0"
            
            .TextMatrix(i, 1) = Val(GetFieldValue(Qry))
            'On Account
            'Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                       "Where (((PAYMENTS.PAYDATE) = #" & .TextMatrix(i, 0) & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
            
            'mySQL old Change these code
            'Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                       "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
            'mySQL
            Qry = "SELECT " & _
                    "SUM(payments.SPAREPARTS) As SumOfSPAREPARTS " & _
                "From " & _
                    "payments " & _
                "Where payments.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND " & _
                    "payments.ORNO IS NOT NULL HAVING Sum(PAYMENTS.SPAREPARTS)<>0"
            
            .TextMatrix(i, 2) = Val(GetFieldValue(Qry))
            'Service
            'Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
                      "Where (((PAYMENTS.PAYDATE) = #" & .TextMatrix(i, 0) & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SERVICE))<>0));"
            'mySQl change these code
            'Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
                      "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.SERVICE))<>0));"
            
            'mySQL
            Qry = "SELECT " & _
                    "Sum(PAYMENTS.SERVICE) AS SumOfSERVICE " & _
                "From " & _
                    "PAYMENTS " & _
                "Where PAYMENTS.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' And " & _
                    "PAYMENTS.ORNO IS NOT NULL HAVING Sum(PAYMENTS.SERVICE)<>0"
            
            .TextMatrix(i, 3) = Val(GetFieldValue(Qry))
            'Down Payment
            'Qry = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT From PAYMENTS " & _
                    "Where (((PAYMENTS.PAYDATE) = #" & .TextMatrix(i, 0) & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.DOWNPAYMENT))<>0));"
            'mySQl change these code
            Qry = "SELECT " & _
                    "Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDOWNPAYMENT " & _
                "From " & _
                    "PAYMENTS " & _
                "Where PAYMENTS.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' And " & _
                    "PAYMENTS.ORNO IS NOT NULL HAVING Sum(PAYMENTS.DOWNPAYMENT)<>0"
            
            .TextMatrix(i, 4) = Val(GetFieldValue(Qry))
            'Others
            'Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
                       "Where (((PAYMENTS.PAYDATE) = #" & .TextMatrix(i, 0) & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
            'mySQl
            Qry = "SELECT " & _
                    "Sum(PAYMENTS.OTHERS) AS SumOfOTHERS " & _
                "From " & _
                    "PAYMENTS " & _
                "Where PAYMENTS.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' And " & _
                    "PAYMENTS.ORNO IS NOT NULL HAVING Sum(PAYMENTS.OTHERS)<>0"
            
            .TextMatrix(i, 5) = Val(GetFieldValue(Qry))
            .TextMatrix(i, 6) = Val(.TextMatrix(i, 1)) + Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5))
            Pbar.Value = Pbar.Value + 1
            DoEvents
            pPct = (Pbar.Value / Pbar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
        Next i
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 6)) > 0 Then
                Qry = "SELECT MSALES.* From MSALES WHERE MSALES.SDATE='" & .TextMatrix(i, 0) & "';"
                If CtrlNumExist(Qry) Then
                    'LYRcon.Execute "UPDATE MSALES SET MSALES.CASH = " & Val(.TextMatrix(i, 1)) & ", MSALES.ACCTCOLL = " & Val(.TextMatrix(i, 2)) & ", MSALES.SRVCCOLL = " & Val(.TextMatrix(i, 3)) & ", MSALES.DP = " & Val(.TextMatrix(i, 4)) & ", MSALES.OTHERS = " & Val(.TextMatrix(i, 5)) & ", MSALES.TTLCOLL = " & Val(.TextMatrix(i, 6)) & _
                                   " WHERE (((MSALES.SDATE)=#" & .TextMatrix(i, 0) & "#));"
                    'mySQl
                    LYRcon.Execute "UPDATE MSALES SET MSALES.CASH = " & Val(.TextMatrix(i, 1)) & ", MSALES.ACCTCOLL = " & Val(.TextMatrix(i, 2)) & ", MSALES.SRVCCOLL = " & Val(.TextMatrix(i, 3)) & ", MSALES.DP = " & Val(.TextMatrix(i, 4)) & ", MSALES.OTHERS = " & Val(.TextMatrix(i, 5)) & ", MSALES.TTLCOLL = " & Val(.TextMatrix(i, 6)) & _
                                   " WHERE MSALES.SDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                    
                Else
                    'Insert Into table Monthly COLLECTION
                    'LYRcon.Execute "INSERT INTO MSALES ( SDATE, CASH, ACCTCOLL, SRVCCOLL, DP, OTHERS, TTLCOLL ) " & _
                                   "VALUES (#" & .TextMatrix(i, 0) & "#, " & Val(.TextMatrix(i, 1)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & Val(.TextMatrix(i, 5)) & ", " & Val(.TextMatrix(i, 6)) & ");"
                    'mySQl
                    LYRcon.Execute "INSERT INTO MSALES ( SDATE, CASH, ACCTCOLL, SRVCCOLL, DP, OTHERS, TTLCOLL ) " & _
                                   "VALUES ('" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 1)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & Val(.TextMatrix(i, 5)) & ", " & Val(.TextMatrix(i, 6)) & ");"
             
                End If
            End If
            Pbar.Value = Pbar.Value + 1
            DoEvents
            pPct = (Pbar.Value / Pbar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
        Next i
    End With
    DoEvents
    
    MsgBox "Finish processing collection for " & MonthName(txtMo.Text) & " " & txtYr.Text & "!", vbInformation + vbOKOnly, "Update Complete"
    Pbar.Visible = False
    lblTotal.Caption = ""
    Frame1.Enabled = True
    cmdTrans(2).Enabled = True
    cmdTrans(3).Enabled = True
    flxSP.Enabled = True
    Call DisplaySales
    
End Sub

