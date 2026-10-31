VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MonthlySalesAdj 
   Caption         =   "Monthly Sales Report"
   ClientHeight    =   9420
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   5655
      Left            =   120
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   3480
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9975
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      BorderStyle     =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxHead 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   2520
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   1720
      _Version        =   393216
      Rows            =   3
      Cols            =   15
      FixedRows       =   3
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      ScrollBars      =   0
      BorderStyle     =   0
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
   Begin VB.Frame Frame2 
      Caption         =   "Sales for the Day"
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
      Height          =   1575
      Left            =   3000
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   840
      Width           =   11415
      Begin MSFlexGridLib.MSFlexGrid flxSDay 
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   4
         Cols            =   13
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         AllowUserResizing=   3
         BorderStyle     =   0
      End
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   8520
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   6960
      TabIndex        =   9
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
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   840
      Width           =   2775
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
         Left            =   240
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblMoColl 
         Alignment       =   2  'Center
         Caption         =   "Total Collection for the Month"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   1080
         Width           =   2535
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1800
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
      Image           =   "MonthlySalesAdj.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3600
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
      Image           =   "MonthlySalesAdj.frx":059A
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPro 
      Height          =   1095
      Left            =   960
      TabIndex        =   12
      Top             =   9720
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      FixedCols       =   0
      FormatString    =   "DATE |>CASH|>ACCTSALES|>SALESRET|>DEBITED|>CREDITED|>SRVCCASH|>SRVCACCT|>DEBITED|>CREDITED|>OTHERS|>TOTALSALES"
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "MonthlySalesAdj.frx":0B34
      cBack           =   -2147483633
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
      TabIndex        =   13
      Tag             =   "XY"
      Top             =   9150
      Width           =   615
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   -120
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14670
      m_MDI_Height    =   9990
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHLY SALES REPORT"
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
      Top             =   480
      Width           =   2610
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
      Caption         =   "MONTHLY SALES REPORT"
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
Attribute VB_Name = "MonthlySalesAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMSALESADJ As New ADODB.Recordset
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
    Dim yrTtl As Currency, Qry As String, MyAcc As String
    Dim dStart As Date, dEnd As Date
    
    Select Case Index
        Case 0  'Reprocess Sales
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=85));"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            Ans = MsgBox("Do you want to Re-Process Sales for then Period " & cboCutOff.Text & "?", vbQuestion + vbYesNo, "Sales")
            If Ans = vbNo Then Exit Sub
            Call ProcessSales
        Case 2  'Print
            dStart = DateSerial(txtYr.Text, 1, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 1, DateSerial(txtYr.Text, txtMo.Text, 1)))
            
            Qry = "SELECT Sum(MSALESADJ.TTLSALES) AS SumOfTTLSALES From MSALESADJ " & _
                  "Where MSALESADJ.SDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' And MSALESADJ.SDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
            
            yrTtl = Val(GetFieldValue(Qry))
            PRNmc.MySource = "MONTHLYSALESADJ"
            PRNmc.QryPrn = txtSQL.Text
            PRNmc.ParamValue1 = "For the Month of " & cboCutOff.Text
                
            
            PRNmc.ParamValue2 = "Total Sales from " & Format(dStart, "Mmmm dd, yyyy") & " to " & Format(dEnd, "Mmmm dd, yyyy") & ": " & Format(yrTtl, "#,##0.00")
            PRNmc.ParamValue3 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='CASHIER';")
            PRNmc.ParamValue4 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='OIC';")
            PRNmc.ParamValue5 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='MANAGER';")
            PRNmc.Show
        Case 3  'Exit
            Unload Me
            MotoRep.tbMain.Visible = True
    End Select
End Sub
Private Sub flxSP_EnterCell()
    ActRow = flxSP.Row
End Sub

Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuINV)
    End If
End Sub
Private Sub Form_Load()
Dim Qry As String, SPcash As Single, SPAcct As Single, SrvCash As Single, SrvAcct As Single, GTtl As Single, Others As Single, SRet As Single
    Dim Mo As Integer, Yr As Integer, MoYr As Boolean, TmpD As Date
    Dim dStart As Date, dEnd As Date, SpDbt As Single, SpCdt As Single, SvcDbt As Single, SvcCdt As Single
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
    
    'Query if the record exist the table
    Qry = "SELECT MSALESADJ.TTLSALES From MSALESADJ WHERE MSALESADJ.TTLSALES>0 AND MSALESADJ.SDATE='" & Format(Date, "yyyy-mm-dd") & "';"
    
    If CtrlNumExist(Qry) Then
        Call LoadTodaySales
    Else
        'Monthly
        dStart = DateSerial(Year(Now), Month(Now), 1)
        dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
         
        'SP Cash
        Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP From payments " & _
              "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL"
        
        SPcash = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(3, 0) = Format(SPcash, "#,##0.00")
        
        'SP Acct
        Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT From SPINVOICE " & _
              "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.INVDATE='" & Format(Date, "yyyy-mm-dd") & "';"
        SPAcct = Val(GetFieldValue(Qry))
        
        flxSDay.TextMatrix(3, 1) = Format(SPAcct, "#,##0.00")
        'SP Sales Ret
        'Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
        '      "WHERE (((SRETURNSP.RETDATE)='" & Format(Date, "yyyy-mm-dd") & "') AND ((SRETURNSP.PROCESS)='REPLACE') AND ((SPINVOICE.INVTYPE)='CASH'));"
        Qry = "SELECT Sum(SRETURNSP.TOTALAMT) As SumOfTOTALAMT From SRETURNSP LEFT OUTER JOIN spinvoice ON (SRETURNSP.INVOICE = spinvoice.SPINVNO) " & _
              "Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
        SRet = Val(GetFieldValue(Qry))
        
        flxSDay.TextMatrix(3, 2) = Format(SRet, "#,##0.00")
        
        
        'SP Insert Debit amount
        Qry = "SELECT SUM(adjdmcm.DEBIT) As SumOfTtlDebit From adjdmcm LEFT OUTER JOIN spinvoice ON (adjdmcm.REFNO = spinvoice.SPINVNO) " & _
              "Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND adjdmcm.DEBIT > 0 AND adjdmcm.REASON <> 'CANCELLED' " & _
              "AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
        SpDbt = Val(GetFieldValue(Qry))
        
        flxSDay.TextMatrix(3, 3) = Format(SpDbt, "#,##0.00")
        'Insert Credit amount
        Qry = "SELECT SUM(adjdmcm.CREDIT) As SumOfTtlCredit From adjdmcm LEFT OUTER JOIN spinvoice ON (adjdmcm.REFNO = spinvoice.SPINVNO) " & _
              "Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND adjdmcm.CREDIT > 0 AND adjdmcm.REASON <> 'CANCELLED' " & _
              "AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
        SpCdt = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(3, 4) = Format(SpCdt, "#,##0.00")
        
        'SP Total
        flxSDay.TextMatrix(3, 5) = Format((((SPcash + SPAcct) - SRet) + SpDbt) - SpCdt, "#,##0.00")
        
        'Service Cash
        Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
              "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') And ((PAYMENTS.CUSTOMERID) = 0)) HAVING (((Sum(PAYMENTS.SERVICE))>0));"
        
        SrvCash = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(3, 6) = Format(SrvCash, "#,##0.00")
           
        Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) AS SumOfTOTALCOST From SRVCJOBORDER " & _
              "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(Date, "yyyy-mm-dd") & "';"
        
        SrvAcct = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(3, 7) = Format(SrvAcct, "#,##0.00")
        
        'Service Insert Debit amount
        Qry = "SELECT SUM(adjdmcm.DEBIT) As SvcSumTtlDebit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
              "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' AND " & _
              "adjdmcm.TRANSTYPE = 'JOB ORDER' AND adjdmcm.REASON <> 'CANCELLED'"
        SvcDbt = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(3, 8) = SvcDbt
        'Service Insert Credit amount
        Qry = "SELECT SUM(adjdmcm.CREDIT) As SvcSumTtlCredit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
              "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' AND " & _
              "adjdmcm.TRANSTYPE = 'JOB ORDER' AND adjdmcm.REASON <> 'CANCELLED'"
        SvcDbt = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(3, 9) = SvcDbt

        'Service Total
        flxSDay.TextMatrix(3, 10) = Format(((SrvCash + SrvAcct) + SvcDbt) - SvcCdt, "#,##0.00")
        Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
                   "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
        
        Others = Val(GetFieldValue(Qry))
        flxSDay.TextMatrix(3, 11) = Format(Others, "#,##0.00")
        'Total
        GTtl = ((((SPcash + SPAcct) - SRet) + SpDbt) - SvcCdt) + ((((SrvCash + SrvAcct) + SvcDbt) - SvcCdt) + Others)
        flxSDay.TextMatrix(3, 12) = Format(GTtl, "#,##0.00")
    End If
    Frame2.Caption = "Sales for the day " & Format(Date, "Dddd, Mmmm dd, yyyy")
    
    Timer1.Enabled = True
    
    MotoRep.sBar.Panels(6).Text = "Monthly Sales Summary"
    MotoRep.tbMain.Visible = False
    ActRow = 1
End Sub
Private Sub LoadTodaySales()
    Dim rsTSales As New ADODB.Recordset
    Dim Qry As String
    Dim TtlSP As Currency, TtlSvc As Currency
    
    Qry = "SELECT MSALESADJ.* From MSALESADJ WHERE MSALESADJ.TTLSALES>0 AND MSALESADJ.SDATE='" & Format(Date, "yyyy-mm-dd") & "';"
    
    With rsTSales
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSDay.TextMatrix(3, 0) = Format(!Cash, "#,##0.00")
            flxSDay.TextMatrix(3, 1) = Format(!AcctSales, "#,##0.00")
            flxSDay.TextMatrix(3, 2) = Format(!SalesRet, "#,##0.00")
            flxSDay.TextMatrix(3, 3) = Format(!ACCTDEBIT, "#,##0.00")
            flxSDay.TextMatrix(3, 4) = Format(!ACCTCREDIT, "#,##0.00")
            TtlSP = (((!Cash + !AcctSales) - !SalesRet) + !ACCTDEBIT) - !ACCTCREDIT
            flxSDay.TextMatrix(3, 5) = Format(TtlSP, "#,##0.00")
            flxSDay.TextMatrix(3, 6) = Format(!SrvcCash, "#,##0.00")
            flxSDay.TextMatrix(3, 7) = Format(!SrvcAcct, "#,##0.00")
            flxSDay.TextMatrix(3, 8) = Format(!SVCDEBIT, "#,##0.00")
            flxSDay.TextMatrix(3, 9) = Format(!SVCCREDIT, "#,##0.00")
            TtlSvc = ((!SrvcCash + !SrvcAcct) + !SVCDEBIT) - !SVCCREDIT
            flxSDay.TextMatrix(3, 10) = Format(TtlSvc, "#,##0.00")
            flxSDay.TextMatrix(3, 11) = Format(!Others, "#,##0.00")
            flxSDay.TextMatrix(3, 12) = Format(TtlSP + TtlSvc + !Others, "#,##0.00")
        End If
        .Close
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsMSALESADJ.State = adStateOpen Then rsMSALESADJ.Close
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, j As Integer
    
    lblHeading(2).Left = Frame2.Left + (Frame2.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame2.Left + (Frame2.Width - lblHeading(3).Width)
    
    With flxHead
        .ColWidth(0) = 150
        For i = 1 To 14
            .ColWidth(i) = (.Width - 250) / 14
            .ColAlignment(i) = 4
        Next i
              
        .TextMatrix(0, 1) = "DATE"
        .TextMatrix(1, 1) = "DATE"
        .TextMatrix(2, 1) = "DATE"
        .TextMatrix(0, 2) = "S P A R E  P A R T S"
        .TextMatrix(0, 3) = "S P A R E  P A R T S"
        .TextMatrix(0, 4) = "S P A R E  P A R T S"
        .TextMatrix(0, 5) = "S P A R E  P A R T S"
        .TextMatrix(0, 6) = "S P A R E  P A R T S"
        .TextMatrix(0, 7) = "S P A R E  P A R T S"
        .TextMatrix(1, 5) = "A D J U S T M E N T"
        .TextMatrix(1, 6) = "A D J U S T M E N T"
        
        .TextMatrix(0, 8) = "S E R V I C E"
        .TextMatrix(0, 9) = "S E R V I C E"
        .TextMatrix(0, 10) = "S E R V I C E"
        .TextMatrix(0, 11) = "S E R V I C E"
        .TextMatrix(0, 12) = "S E R V I C E"
        .TextMatrix(1, 10) = "A D J U S T M E N T"
        .TextMatrix(1, 11) = "A D J U S T M E N T"
 
        .TextMatrix(1, 2) = "CASH"
        .TextMatrix(2, 2) = "CASH"
        .TextMatrix(1, 3) = "ON-ACCT."
        .TextMatrix(2, 3) = "ON-ACCT."
        .TextMatrix(1, 4) = "SALES RET."
        .TextMatrix(2, 4) = "SALES RET."
        .TextMatrix(2, 5) = "DEBIT"
        .TextMatrix(2, 6) = "CREDIT"
        .TextMatrix(1, 7) = "TOTAL"
        .TextMatrix(2, 7) = "TOTAL"
          
        .TextMatrix(1, 8) = "CASH"
        .TextMatrix(2, 8) = "CASH"
        .TextMatrix(1, 9) = "ON-ACCT."
        .TextMatrix(2, 9) = "ON-ACCT."
        .TextMatrix(2, 10) = "DEBIT"
        .TextMatrix(2, 11) = "CREDIT"
        .TextMatrix(1, 12) = "TOTAL"
        .TextMatrix(2, 12) = "TOTAL"
        .TextMatrix(0, 13) = "OTHERS"
        .TextMatrix(1, 13) = "OTHERS"
        .TextMatrix(2, 13) = "OTHERS"
        .TextMatrix(0, 14) = "GRAND"
        .TextMatrix(1, 14) = "TOTAL"
        .TextMatrix(2, 14) = "TOTAL"
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
        .MergeCol(9) = True
        .MergeCol(10) = True
        .MergeCol(12) = True
        .MergeCol(13) = True
        .MergeCol(14) = True
    End With
    With flxSP
        .Top = flxHead.Top + flxHead.Height
        .ColWidth(0) = 150
        For i = 1 To 13
            .ColWidth(i) = (.Width - 250) / 14
            If i = 1 Then
                .ColAlignment(i) = 4
            Else
                .ColAlignment(i) = 7
            End If
        Next i
    End With
    With flxSDay
        For i = 0 To 12
            .ColWidth(i) = (.Width) / 13
            .ColAlignment(i) = 4
        Next i
    
        .TextMatrix(0, 0) = "S P A R E  P A R T S"
        .TextMatrix(0, 1) = "S P A R E  P A R T S"
        .TextMatrix(0, 2) = "S P A R E  P A R T S"
        .TextMatrix(0, 3) = "S P A R E  P A R T S"
        .TextMatrix(0, 4) = "S P A R E  P A R T S"
        .TextMatrix(0, 5) = "S P A R E  P A R T S"
        .TextMatrix(1, 3) = "ADJUSTMENT"
        .TextMatrix(1, 4) = "ADJUSTMENT"
        .TextMatrix(0, 6) = "S E R V I C E"
        .TextMatrix(0, 7) = "S E R V I C E"
        .TextMatrix(0, 8) = "S E R V I C E"
        .TextMatrix(0, 9) = "S E R V I C E"
        .TextMatrix(0, 10) = "S E R V I C E"
        
        .TextMatrix(1, 0) = "CASH"
        .TextMatrix(2, 0) = "CASH"
        .TextMatrix(1, 1) = "ON-ACCT."
        .TextMatrix(2, 1) = "ON-ACCT."
        .TextMatrix(1, 2) = "SALES RET."
        .TextMatrix(2, 2) = "SALES RET."
        .TextMatrix(2, 3) = "DEBIT"
        .TextMatrix(2, 4) = "CREDIT"
        .TextMatrix(1, 5) = "TOTAL"
        .TextMatrix(2, 5) = "TOTAL"
        
        
        .TextMatrix(1, 6) = "CASH"
        .TextMatrix(2, 6) = "CASH"
        .TextMatrix(1, 7) = "ON-ACCT."
        .TextMatrix(2, 7) = "ON-ACCT."
        .TextMatrix(2, 8) = "DEBIT"
        .TextMatrix(2, 9) = "CREDIT"
        .TextMatrix(1, 10) = "TOTAL"
        .TextMatrix(2, 10) = "TOTAL"
        .TextMatrix(0, 11) = "OTHERS"
        .TextMatrix(1, 11) = "OTHERS"
        .TextMatrix(2, 11) = "OTHERS"
        .TextMatrix(0, 12) = "GRAND"
        .TextMatrix(1, 12) = "TOTAL"
        .TextMatrix(2, 12) = "TOTAL"
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
        .MergeCol(9) = True
        .MergeCol(10) = True
        .MergeCol(11) = True
        .MergeCol(12) = True
        
     
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
    'Qry = "SELECT MSALESADJ.* From MSALESADJ Where (((MSALESADJ.SDATE) >= #" & dStart & "# And (MSALESADJ.SDATE) <= #" & dEnd & "#) And ((MSALESADJ.TTLSALES) > 0)) ORDER BY MSALESADJ.SDATE;"
    'txtSQL.Text = "SELECT MSALESADJ.SDATE, MSALESADJ.CASH, MSALESADJ.ACCTSALES, MSALESADJ.SALESRET, MSALESADJ.SRVCCASH, MSALESADJ.SRVCACCT, MSALESADJ.OTHERS From MSALESADJ " & _
                  "Where (((MSALESADJ.SDATE) >= #" & dStart & "# And (MSALESADJ.SDATE) <= #" & dEnd & "#) And ((MSALESADJ.TTLSALES) > 0)) ORDER BY MSALESADJ.SDATE;"
    'mySQL
    Qry = "SELECT MSALESADJ.* From MSALESADJ Where MSALESADJ.SDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' And MSALESADJ.SDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' And MSALESADJ.TTLSALES > 0 ORDER BY MSALESADJ.SDATE;"
    txtSQL.Text = "SELECT MSALESADJ.SDATE, MSALESADJ.CASH, MSALESADJ.ACCTSALES, MSALESADJ.SALESRET, MSALESADJ.ACCTDEBIT, MSALESADJ.ACCTCREDIT, MSALESADJ.SRVCCASH, MSALESADJ.SVCDEBIT, MSALESADJ.SVCCREDIT, MSALESADJ.SRVCACCT, MSALESADJ.OTHERS " & _
                  "From MSALESADJ Where (((MSALESADJ.SDATE) >= '" & Format(dStart, "yyyy-mm-dd") & "' And (MSALESADJ.SDATE) <= '" & Format(dEnd, "yyyy-mm-dd") & "') And ((MSALESADJ.TTLSALES) > 0)) ORDER BY MSALESADJ.SDATE;"
    
    
    If rsMSALESADJ.State = adStateOpen Then rsMSALESADJ.Close
    rsMSALESADJ.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsMSALESADJ.RecordCount = 0 Then
        flxSP.Rows = 0
        Ans = MsgBox("Do you want to Process Sales for then Period " & cboCutOff.Text & "?", vbQuestion + vbYesNo, "Sales")
        If Ans = vbNo Then Exit Sub
        Call ProcessSales
        Exit Sub
    End If
    With rsMSALESADJ
        .MoveFirst
        flxSP.Rows = 0
        Do While Not .EOF
            'Old
            'TtlSP = (!Cash + !AcctSales) - !SalesRet
            'TtlSrv = !SrvcCash + !SrvcAcct
            'GTtl = TtlSP + TtlSrv + !Others
            
            TtlSP = (((!Cash + !AcctSales) - !SalesRet) + !ACCTDEBIT) - !ACCTCREDIT
            TtlSrv = ((!SrvcCash + !SrvcAcct) + !SVCDEBIT) - !SVCCREDIT
            GTtl = TtlSP + TtlSrv + !Others
            
            
            'Insert DM/CM here
            'flxSP.AddItem vbTab & Format(!SDATE, "mm/dd/yyyy") & vbTab & Format(!Cash, "#,##0.00") & vbTab & Format(!AcctSales, "#,##0.00") & vbTab & Format(!SalesRet, "#,##0.00") & vbTab & Format(TtlSP, "#,##0.00") & vbTab & Format(!SrvcCash, "#,##0.00") & vbTab & Format(!SrvcAcct, "#,##0.00") & vbTab & Format(TtlSrv, "#,##0.00") & vbTab & Format(!Others, "#,##0.00") & vbTab & Format(GTtl, "#,##0.00")
            flxSP.AddItem vbTab & Format(!SDATE, "mm/dd/yyyy") & vbTab & Format(!Cash, "#,##0.00") & vbTab & Format(!AcctSales, "#,##0.00") & vbTab & Format(!SalesRet, "#,##0.00") & vbTab & Format(!ACCTDEBIT, "#,##0.00") & vbTab & Format(!ACCTCREDIT, "#,##0.00") & vbTab & Format(TtlSP, "#,##0.00") & vbTab & Format(!SrvcCash, "#,##0.00") & vbTab & Format(!SrvcAcct, "#,##0.00") & vbTab & Format(!SVCDEBIT, "#,##0.00") & vbTab & Format(!SVCCREDIT, "#,##0.00") & vbTab & Format(TtlSrv, "#,##0.00") & vbTab & Format(!Others, "#,##0.00") & vbTab & Format(GTtl, "#,##0.00")
            
            .MoveNext
        Loop
        .Close
    End With
    ActRow = 1
    
    Qry = "SELECT Sum(MSALESADJ.TTLSALES) AS SumOfTTLSALES From MSALESADJ WHERE MSALESADJ.SDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And MSALESADJ.SDATE<='" & Format(dEnd, "yyyy-mm-dd") & "';"
    lblMoColl.Caption = "Total Sales for the Month: " & Format(Val(GetFieldValue(Qry)), "#,##0.00")
End Sub
Private Sub ProcessSales()
    Dim Qry As String, i As Integer, Ans As Integer, pPct As Single, SvcDisc As Currency, SvcAcct As Currency
    Dim dStart As Date, dEnd As Date, TmpD As Date
    
    
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
            'SP Cash
            Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                     "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') And (Not (PAYMENTS.INVOICENO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
            
            .TextMatrix(i, 1) = Val(GetFieldValue(Qry))
            'SP On Account
            Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT From SPINVOICE " & _
                  "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.INVDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                        
            .TextMatrix(i, 2) = Val(GetFieldValue(Qry))
            
            'Sales Return base on return date
            'Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                  "WHERE (((SRETURNSP.RETDATE)='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "'));"
            
            'Sales return base on invoice
            Qry = "SELECT Sum(SRETURNSP.TOTALAMT) As SumOfTOTALAMT From SRETURNSP LEFT OUTER JOIN spinvoice ON (SRETURNSP.INVOICE = spinvoice.SPINVNO) " & _
                  "Where spinvoice.INVDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
            .TextMatrix(i, 3) = Val(GetFieldValue(Qry))
            
            'SP Insert Debit amount
            Qry = "SELECT SUM(adjdmcm.DEBIT) As SumOfTtlDebit From adjdmcm LEFT OUTER JOIN spinvoice ON (adjdmcm.REFNO = spinvoice.SPINVNO) " & _
                  "Where spinvoice.INVDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND adjdmcm.DEBIT > 0 AND adjdmcm.REASON <> 'CANCELLED' " & _
                  "AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
            .TextMatrix(i, 4) = Val(GetFieldValue(Qry))
            
            'Insert Credit amount
            Qry = "SELECT SUM(adjdmcm.CREDIT) As SumOfTtlCredit From adjdmcm LEFT OUTER JOIN spinvoice ON (adjdmcm.REFNO = spinvoice.SPINVNO) " & _
                  "Where spinvoice.INVDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND adjdmcm.CREDIT > 0 AND adjdmcm.REASON <> 'CANCELLED' " & _
                  "AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
      
            .TextMatrix(i, 5) = Val(GetFieldValue(Qry))
            
            'Service Cash
            Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
                  "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') And ((PAYMENTS.CUSTOMERID) = 0)) HAVING (((Sum(PAYMENTS.SERVICE))>0));"
            .TextMatrix(i, 6) = Val(GetFieldValue(Qry))
                        
            Qry = "SELECT SUM(adjdmcm.DEBIT) As SumSvcTtlDebit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
                  "Where payments.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND  payments.CUSTOMERID = 0 AND payments.SERVICE > 0"
            
            'Service Credit
            Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) AS SumOfTOTALCOST From SRVCJOBORDER " & _
                  "WHERE (((SRVCJOBORDER.PAYTYPE)='ON ACCOUNT') AND ((SRVCJOBORDER.DATEPROMISED)='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') AND ((SRVCJOBORDER.PRINTCOPY)>0));"
            SvcAcct = Val(GetFieldValue(Qry))
            
            Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST*(SRVCJOBORDER.DISCOUNT/100)) AS DISCAMT From SRVCJOBORDER " & _
                  "WHERE (((SRVCJOBORDER.DISCOUNT)>0) AND ((SRVCJOBORDER.PAYTYPE)='ON ACCOUNT') AND ((SRVCJOBORDER.DATEPROMISED)='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') AND ((SRVCJOBORDER.PRINTCOPY)>0));"
            
            'Service on servive account less service discount
            SvcDisc = Val(GetFieldValue(Qry))
            .TextMatrix(i, 7) = SvcAcct - SvcDisc
            
            'Service Insert Debit amount
            Qry = "SELECT SUM(adjdmcm.DEBIT) As SvcSumTtlDebit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
                  "Where payments.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' AND " & _
                  "adjdmcm.TRANSTYPE = 'JOB ORDER' AND adjdmcm.REASON <> 'CANCELLED'"
            .TextMatrix(i, 8) = Val(GetFieldValue(Qry))
            
            Qry = "SELECT SUM(adjdmcm.CREDIT) As SvcSumTtlCredit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
                  "Where payments.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' AND " & _
                  "adjdmcm.TRANSTYPE = 'JOB ORDER' AND adjdmcm.REASON <> 'CANCELLED'"
            .TextMatrix(i, 9) = Val(GetFieldValue(Qry))
            'Others
            Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
                       "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
            
            .TextMatrix(i, 10) = Val(GetFieldValue(Qry))
           
            .TextMatrix(i, 11) = (Val(.TextMatrix(i, 1)) + Val(.TextMatrix(i, 2))) - Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6))
            Pbar.Value = Pbar.Value + 1
            DoEvents
            pPct = (Pbar.Value / Pbar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
        Next i
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 7)) > 0 Then
                Qry = "SELECT MSALESADJ.* From MSALESADJ WHERE MSALESADJ.SDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                'Update table here dc/cm for customer account
                If CtrlNumExist(Qry) Then
                     LYRcon.Execute "UPDATE MSALESADJ SET MSALESADJ.CASH = " & Val(.TextMatrix(i, 1)) & ", MSALESADJ.ACCTSALES = " & Val(.TextMatrix(i, 2)) & ", MSALESADJ.SALESRET = " & Val(.TextMatrix(i, 3)) & ", MSALESADJ.ACCTDEBIT = " & Val(.TextMatrix(i, 4)) & ",MSALESADJ.ACCTCREDIT = " & Val(.TextMatrix(i, 5)) & ", MSALESADJ.SRVCCASH = " & Val(.TextMatrix(i, 6)) & ", MSALESADJ.SRVCACCT = " & Val(.TextMatrix(i, 7)) & ", MSALESADJ.SVCDEBIT = " & Val(.TextMatrix(i, 8)) & ", MSALESADJ.SVCCREDIT = " & Val(.TextMatrix(i, 9)) & ", MSALESADJ.OTHERS = " & Val(.TextMatrix(i, 10)) & ", MSALESADJ.TTLSALES = " & Val(.TextMatrix(i, 11)) & _
                                   " WHERE MSALESADJ.SDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                Else
                    'Insert Into table Monthly Sales
                    LYRcon.Execute "INSERT INTO MSALESADJ ( SDATE, CASH, ACCTSALES, SALESRET, ACCTDEBIT, ACCTCREDIT, SRVCCASH, SRVCACCT, SVCDEBIT, SVCCREDIT, OTHERS, TTLSALES ) " & _
                                   "VALUES ('" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 1)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(.TextMatrix(i, 3)) & "," & Val(.TextMatrix(i, 4)) & "," & Val(.TextMatrix(i, 5)) & ", " & Val(.TextMatrix(i, 6)) & ", " & Val(.TextMatrix(i, 7)) & ", " & Val(.TextMatrix(i, 8)) & ", " & Val(.TextMatrix(i, 9)) & ", " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 11)) & ");"
                End If
            End If
            Pbar.Value = Pbar.Value + 1
            DoEvents
            pPct = (Pbar.Value / Pbar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
        Next i
    End With
    DoEvents
    
    MsgBox "Finish processing Sales for " & MonthName(txtMo.Text) & " " & txtYr.Text & "!", vbInformation + vbOKOnly, "Update Complete"
    Pbar.Visible = False
    lblTotal.Caption = ""
    Frame1.Enabled = True
    cmdTrans(2).Enabled = True
    cmdTrans(3).Enabled = True
    flxSP.Enabled = True
    Call DisplaySales
    
End Sub



