VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MonthlySales 
   Caption         =   "Monthly Sales Report"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   6375
      Left            =   120
      TabIndex        =   19
      Tag             =   "XYF"
      Top             =   2760
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      BorderStyle     =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxHead 
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Tag             =   "XYF"
      Top             =   2160
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   1085
      _Version        =   393216
      Cols            =   11
      FixedRows       =   2
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
      Height          =   1335
      Left            =   4680
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   720
      Width           =   9615
      Begin MSFlexGridLib.MSFlexGrid flxSDay 
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   285
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   3
         Cols            =   9
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   315
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
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
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   4455
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
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   172
         Width           =   2655
      End
      Begin VB.Label lblMoColl 
         Alignment       =   2  'Center
         Caption         =   "Total Collection for the Month"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   600
         Width           =   4215
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
      Image           =   "MonthlySales.frx":0000
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
      Image           =   "MonthlySales.frx":059A
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPro 
      Height          =   1095
      Left            =   960
      TabIndex        =   12
      Top             =   9720
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      FormatString    =   "DATE |>CASH|>ACCTSALES|>SALESRET|>SRVCCASH|>SRVCACCT|>OTHERS|>TOTALSALES"
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   20
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
      Image           =   "MonthlySales.frx":0B34
      Enabled         =   0   'False
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
      Left            =   0
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
Attribute VB_Name = "MonthlySales"
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
    Dim yrTtl As Currency, Qry As String, MyAcc As String
    Dim dStart As Date, dEnd As Date
    
    Select Case Index
        Case 0  'Reprocess Sales
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=85));"
            'mySQl
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
            dStart = DateSerial(txtYr.Text, txtMo.Text, 1)
            dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
            'Qry = "SELECT Sum(MSALES.TTLSALES) AS SumOfTTLSALES From MSALES " & _
                  "Where MSALES.SDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' And MSALES.SDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
            'mySQl
            Qry = "call sp_mrcretails.sp_dailySalesReport_perMonthRange('" & sysBranchCode & "','" & Format(dStart, "yyyy-mm-dd") & "','" & Format(dEnd, "yyyy-mm-dd") & "')"
    
            PRNmc.MySource = "MONTHLYSALES"
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = SysBranch
            PRNmc.ParamValue2 = SysAddress
            PRNmc.ParamValue3 = "For the Month of " & cboCutOff.Text
            PRNmc.ParamValue4 = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEENAME From Employees Where EMPID = " & Val(UserID) & " ")
            PRNmc.ParamValue5 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='MANAGER';")
            PRNmc.ParamValue6 = GetFieldValue("SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='OIC';")
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


Private Sub Form_Load()
    Dim Qry As String, SPcash As Single, SPAcct As Single, SrvCash As Single, SrvAcct As Single, GTtl As Single, Others As Single, SRet As Single
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
        
    Frame2.Caption = "Sales for the day " & Format(Date, "Dddd, Mmmm dd, yyyy")
    
    Timer1.Enabled = True
    
    MotoRep.sBar.Panels(6).Text = "Monthly Sales Summary"
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
    
    lblHeading(2).Left = Frame2.Left + (Frame2.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame2.Left + (Frame2.Width - lblHeading(3).Width)
    
    With flxHead
        .ColWidth(0) = 300
        For i = 1 To 10
            .ColWidth(i) = (.Width - 400) / 10
            .ColAlignment(i) = 4
        Next i
        .TextMatrix(0, 1) = "DATE"
        .TextMatrix(0, 2) = "S P A R E  P A R T S"
        .TextMatrix(0, 3) = "S P A R E  P A R T S"
        .TextMatrix(0, 4) = "S P A R E  P A R T S"
        .TextMatrix(0, 5) = "S P A R E  P A R T S"
        .TextMatrix(0, 6) = "S E R V I C E"
        .TextMatrix(0, 7) = "S E R V I C E"
        .TextMatrix(0, 8) = "S E R V I C E"
        .TextMatrix(0, 9) = "OTHERS"
        .TextMatrix(0, 10) = "GRAND TOTAL"
        .TextMatrix(1, 1) = "DATE"
        .TextMatrix(1, 2) = "CASH"
        .TextMatrix(1, 3) = "ON-ACCT."
        .TextMatrix(1, 4) = "SALES RET."
        .TextMatrix(1, 5) = "TOTAL"
        .TextMatrix(1, 6) = "CASH"
        .TextMatrix(1, 7) = "ON-ACCT."
        .TextMatrix(1, 8) = "TOTAL"
        .TextMatrix(1, 9) = "OTHERS"
        .TextMatrix(1, 10) = "GRAND TOTAL"
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(9) = True
        .MergeCol(10) = True
    End With
    With flxSP
        .Top = flxHead.Top + flxHead.Height
        .ColWidth(0) = 300
        For i = 1 To 10
            .ColWidth(i) = (.Width - 400) / 10
            If i = 1 Then
                .ColAlignment(i) = 4
            Else
                .ColAlignment(i) = 7
            End If
        Next i
    End With
    With flxSDay
        For i = 0 To 8
            .ColWidth(i) = (.Width) / 9
            .ColAlignment(i) = 4
        Next i
        .TextMatrix(0, 0) = "S P A R E  P A R T S"
        .TextMatrix(0, 1) = "S P A R E  P A R T S"
        .TextMatrix(0, 2) = "S P A R E  P A R T S"
        .TextMatrix(0, 3) = "S P A R E  P A R T S"
        .TextMatrix(0, 4) = "S E R V I C E"
        .TextMatrix(0, 5) = "S E R V I C E"
        .TextMatrix(0, 6) = "S E R V I C E"
        .TextMatrix(0, 7) = "OTHERS"
        .TextMatrix(0, 8) = "T O T A L"
        .TextMatrix(1, 0) = "CASH"
        .TextMatrix(1, 1) = "ON-ACCT"
        .TextMatrix(1, 2) = "SAL. RET."
        .TextMatrix(1, 3) = "TOTAL"
        .TextMatrix(1, 4) = "CASH"
        .TextMatrix(1, 5) = "ON-ACCT"
        .TextMatrix(1, 6) = "TOTAL"
        .TextMatrix(1, 7) = "OTHERS"
        .TextMatrix(1, 8) = "T O T A L"
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
    End With
End Sub

Private Sub Timer1_Timer()
    cboCutOff.ListIndex = 0
    Call LoadTodaySales
    Call DisplaySales
    Timer1.Enabled = False
End Sub

Public Sub LoadTodaySales()
Dim Qry As String

    Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP From payments " & _
          "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL"

    SPcash = Val(GetFieldValue(Qry))
    flxSDay.TextMatrix(2, 0) = Format(SPcash, "#,##0.00")
    
    'SP Acct
    Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT From SPINVOICE " & _
          "WHERE (((SPINVOICE.INVTYPE)='CREDIT') AND ((SPINVOICE.INVDATE)='" & Format(Date, "yyyy-mm-dd") & "'));"

    SPAcct = Val(GetFieldValue(Qry))
    flxSDay.TextMatrix(2, 1) = Format(SPAcct, "#,##0.00")

    Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE (((SRETURNSP.RETDATE)='" & Format(Date, "yyyy-mm-dd") & "') AND ((SRETURNSP.PROCESS)='REPLACE') AND ((SPINVOICE.INVTYPE)='CASH'));"

    SRet = Val(GetFieldValue(Qry))
    flxSDay.TextMatrix(2, 2) = Format(SRet, "#,##0.00")
    
    'SP Total
    flxSDay.TextMatrix(2, 3) = Format((SPcash + SPAcct) - SRet, "#,##0.00")
    
    'Service Cash
    Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
          "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') And ((PAYMENTS.CUSTOMERID) = 0)) HAVING (((Sum(PAYMENTS.SERVICE))>0));"

    SrvCash = Val(GetFieldValue(Qry))
    flxSDay.TextMatrix(2, 4) = Format(SrvCash, "#,##0.00")
    
    'Service On-Account
    Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) AS SumOfTOTALCOST From SRVCJOBORDER " & _
          "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(Date, "yyyy-mm-dd") & "';"

    SrvAcct = Val(GetFieldValue(Qry))
    flxSDay.TextMatrix(2, 5) = Format(SrvAcct, "#,##0.00")
    
    'Service Total
    flxSDay.TextMatrix(2, 6) = Format(SrvCash + SrvAcct, "#,##0.00")

    'Others
    Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
               "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"

    Others = Val(GetFieldValue(Qry))
    flxSDay.TextMatrix(2, 7) = Format(Others, "#,##0.00")
    'Total
    GTtl = ((SPcash + SPAcct) - SRet) + SrvCash + SrvAcct + Others
    flxSDay.TextMatrix(2, 8) = Format(GTtl, "#,##0.00")

End Sub

Private Sub DisplaySales()
    Dim Qry As String
    Dim dStart As Date, dEnd As Date
    Dim TtlSP As Currency, TtlSrv As Currency, GTtl As Currency
    
    dStart = DateSerial(txtYr.Text, txtMo.Text, 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    'Qry = "SELECT MSALES.* From MSALES Where (((MSALES.SDATE) >= #" & dStart & "# And (MSALES.SDATE) <= #" & dEnd & "#) And ((MSALES.TTLSALES) > 0)) ORDER BY MSALES.SDATE;"
    'txtSQL.Text = "SELECT MSALES.SDATE, MSALES.CASH, MSALES.ACCTSALES, MSALES.SALESRET, MSALES.SRVCCASH, MSALES.SRVCACCT, MSALES.OTHERS From MSALES " & _
                  "Where (((MSALES.SDATE) >= #" & dStart & "# And (MSALES.SDATE) <= #" & dEnd & "#) And ((MSALES.TTLSALES) > 0)) ORDER BY MSALES.SDATE;"
    'mySQL
'    Qry = "SELECT MSALES.* From MSALES Where MSALES.SDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' And MSALES.SDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' And MSALES.TTLSALES > 0 ORDER BY MSALES.SDATE;"
'    txtSQL.Text = "SELECT MSALES.SDATE, MSALES.CASH, MSALES.ACCTSALES, MSALES.SALESRET, MSALES.SRVCCASH, MSALES.SRVCACCT, MSALES.OTHERS From MSALES " & _
'                  "Where (((MSALES.SDATE) >= '" & Format(dStart, "yyyy-mm-dd") & "' And (MSALES.SDATE) <= '" & Format(dEnd, "yyyy-mm-dd") & "') And ((MSALES.TTLSALES) > 0)) ORDER BY MSALES.SDATE;"
'
'    If rsMSales.State = adStateOpen Then rsMSales.Close
'    rsMSales.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'    If rsMSales.RecordCount = 0 Then
'        flxSP.Rows = 0
'        Ans = MsgBox("Do you want to Process Sales for then Period " & cboCutOff.Text & "?", vbQuestion + vbYesNo, "Sales")
'        If Ans = vbNo Then Exit Sub
'        Call ProcessSales
'        Exit Sub
'    End If
'    With rsMSales
'        .MoveFirst
'        flxSP.Rows = 0
'        Do While Not .EOF
'            TtlSP = (!Cash + !AcctSales) - !SalesRet    'To be continued ...
'            TtlSrv = !SrvcCash + !SrvcAcct
'            GTtl = TtlSP + TtlSrv + !Others
'            flxSP.AddItem vbTab & Format(!SDATE, "mm/dd/yyyy") & vbTab & Format(!Cash, "#,##0.00") & vbTab & Format(!AcctSales, "#,##0.00") & vbTab & Format(!SalesRet, "#,##0.00") & vbTab & Format(TtlSP, "#,##0.00") & vbTab & Format(!SrvcCash, "#,##0.00") & vbTab & Format(!SrvcAcct, "#,##0.00") & vbTab & Format(TtlSrv, "#,##0.00") & vbTab & Format(!Others, "#,##0.00") & vbTab & Format(GTtl, "#,##0.00")
'            .MoveNext
'        Loop
'        .Close
'    End With
'    ActRow = 1
'    'Qry = "SELECT Sum(MSALES.TTLSALES) AS SumOfTTLSALES From MSALES WHERE (((MSALES.SDATE)>=#" & dStart & "# And (MSALES.SDATE)<=#" & dEnd & "#));"
'
'    'mySQL
'    Qry = "SELECT Sum(MSALES.TTLSALES) AS SumOfTTLSALES From MSALES WHERE MSALES.SDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And MSALES.SDATE<='" & Format(dEnd, "yyyy-mm-dd") & "';"
'

    Qry = "call sp_mrcretails.sp_dailySalesReport_perMonthRange('" & sysBranchCode & "','" & Format(dStart, "yyyy-mm-dd") & "','" & Format(dEnd, "yyyy-mm-dd") & "')"
    
    With rsMSales
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then
            flxSP.Rows = 0
            Exit Sub
        End If
        
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 0
            GTtl = 0
            flxSP.Rows = .RecordCount
            For i = 0 To flxSP.Rows - 1
                'TRANSDATE, SPCASH, SPCREDIT, SPRETURN, SPTOTAL, SRVCCASH, SRVCCREDIT, SRVCTOTAL, OTHERS, GRANDTOTAL
                flxSP.TextMatrix(i, 1) = Format(!TRANSDATE, "mm/dd/yyyy")
                flxSP.TextMatrix(i, 2) = Format(!SPcash, "#,###,##0.00")
                flxSP.TextMatrix(i, 3) = Format(!SPCREDIT, "#,###,##0.00")
                flxSP.TextMatrix(i, 4) = Format(!SPRETURN, "#,###,##0.00")
                flxSP.TextMatrix(i, 5) = Format(!SPTOTAL, "#,###,##0.00")
                flxSP.TextMatrix(i, 6) = Format(!SRVCCASH, "#,###,##0.00")
                flxSP.TextMatrix(i, 7) = Format(!SRVCCREDIT, "#,###,##0.00")
                flxSP.TextMatrix(i, 8) = Format(!SRVCTOTAL, "#,###,##0.00")
                flxSP.TextMatrix(i, 9) = Format(!Others, "#,###,##0.00")
                flxSP.TextMatrix(i, 10) = Format(!GRANDTOTAL, "#,###,##0.00")
                GTtl = GTtl + Format(!GRANDTOTAL, "#,###,##0.00")
                'flxSP.AddItem vbTab & Format(!SDATE, "mm/dd/yyyy") & vbTab & Format(!Cash, "#,##0.00") & vbTab & Format(!AcctSales, "#,##0.00") & vbTab & Format(!SalesRet, "#,##0.00") & vbTab & Format(TtlSP, "#,##0.00") & vbTab & Format(!SRVCCASH, "#,##0.00") & vbTab & Format(!SrvcAcct, "#,##0.00") & vbTab & Format(TtlSrv, "#,##0.00") & vbTab & Format(!Others, "#,##0.00") & vbTab & Format(GTtl, "#,##0.00")
                .MoveNext
            Next i
            .Close
        End If
    End With
    lblMoColl.Caption = "Total Sales for the Month: " & Format(GTtl, "#,###,##0.00")
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
            'Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                     "Where (((PAYMENTS.PAYDATE) = #" & .TextMatrix(i, 0) & "#) And (Not (PAYMENTS.INVOICENO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
            'mySQl
            Qry = "SELECT Sum(PAYMENTS.SPAREPARTS) AS SumOfSPAREPARTS From PAYMENTS " & _
                     "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') And (Not (PAYMENTS.INVOICENO) Is Null)) HAVING (((Sum(PAYMENTS.SPAREPARTS))<>0));"
            
            .TextMatrix(i, 1) = Val(GetFieldValue(Qry))
            'SP On Account
            'Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT From SPINVOICE " & _
                  "WHERE (((SPINVOICE.INVTYPE)='CREDIT') AND ((SPINVOICE.INVDATE)=#" & .TextMatrix(i, 0) & "#));"
            'mySQl
            'Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT From SPINVOICE " & _
                  "WHERE (((SPINVOICE.INVTYPE='CREDIT') AND ((SPINVOICE.INVDATE)='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "'));"
            Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT From SPINVOICE " & _
                  "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.INVDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                        
            .TextMatrix(i, 2) = Val(GetFieldValue(Qry))
            
            'Sales Return
            'Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                  "WHERE (((SRETURNSP.RETDATE)=#" & .TextMatrix(i, 0) & "#));"
            'mySQL
            Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                  "WHERE (((SRETURNSP.RETDATE)='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "'));"
            
            .TextMatrix(i, 3) = Val(GetFieldValue(Qry))
            'Service Cash
            'Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
                  "Where (((PAYMENTS.PAYDATE) = #" & .TextMatrix(i, 0) & "#) And ((PAYMENTS.CUSTOMERID) = 0)) HAVING (((Sum(PAYMENTS.SERVICE))>0));"
            'mySQL
            Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE From PAYMENTS " & _
                  "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') And ((PAYMENTS.CUSTOMERID) = 0)) HAVING (((Sum(PAYMENTS.SERVICE))>0));"
            
            .TextMatrix(i, 4) = Val(GetFieldValue(Qry))
            'Service Credit
            'Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) AS SumOfTOTALCOST From SRVCJOBORDER " & _
                  "WHERE (((SRVCJOBORDER.PAYTYPE)='ON ACCOUNT') AND ((SRVCJOBORDER.DATEPROMISED)=#" & .TextMatrix(i, 0) & "#) AND ((SRVCJOBORDER.PRINTCOPY)>0));"
            'mySQL
            
            Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) AS SumOfTOTALCOST From SRVCJOBORDER " & _
                  "WHERE (((SRVCJOBORDER.PAYTYPE)='ON ACCOUNT') AND ((SRVCJOBORDER.DATEPROMISED)='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') AND ((SRVCJOBORDER.PRINTCOPY)>0));"
            SvcAcct = Val(GetFieldValue(Qry))
            'Qry = "SELECT Sum([TOTALCOST]*([DISCOUNT]/100)) AS DISCAMT From SRVCJOBORDER " & _
                  "WHERE (((SRVCJOBORDER.DISCOUNT)>0) AND ((SRVCJOBORDER.PAYTYPE)='ON ACCOUNT') AND ((SRVCJOBORDER.DATEPROMISED)=#" & .TextMatrix(i, 0) & "#) AND ((SRVCJOBORDER.PRINTCOPY)>0));"
            'mySQl
            Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST*(SRVCJOBORDER.DISCOUNT/100)) AS DISCAMT From SRVCJOBORDER " & _
                  "WHERE (((SRVCJOBORDER.DISCOUNT)>0) AND ((SRVCJOBORDER.PAYTYPE)='ON ACCOUNT') AND ((SRVCJOBORDER.DATEPROMISED)='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') AND ((SRVCJOBORDER.PRINTCOPY)>0));"
            
            SvcDisc = Val(GetFieldValue(Qry))
            .TextMatrix(i, 5) = SvcAcct - SvcDisc
            'Others
            'Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
                       "Where (((PAYMENTS.PAYDATE) = #" & .TextMatrix(i, 0) & "#) And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
            'mySQl
            Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS From PAYMENTS " & _
                       "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') And (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
            
            .TextMatrix(i, 6) = Val(GetFieldValue(Qry))
            
            .TextMatrix(i, 7) = (Val(.TextMatrix(i, 1)) + Val(.TextMatrix(i, 2))) - Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6))
            Pbar.Value = Pbar.Value + 1
            DoEvents
            pPct = (Pbar.Value / Pbar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
        Next i

        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 7)) > 0 Then
                'Qry = "SELECT MSALES.* From MSALES WHERE (((MSALES.SDATE)=#" & .TextMatrix(i, 0) & "#));"
                'mySQL
                Qry = "SELECT MSALES.* From MSALES WHERE MSALES.SDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                
                If CtrlNumExist(Qry) Then
                    'LYRcon.Execute "UPDATE MSALES SET MSALES.CASH = " & Val(.TextMatrix(i, 1)) & ", MSALES.ACCTSALES = " & Val(.TextMatrix(i, 2)) & ", MSALES.SALESRET = " & Val(.TextMatrix(i, 3)) & ", MSALES.SRVCCASH = " & Val(.TextMatrix(i, 4)) & ", MSALES.SRVCACCT = " & Val(.TextMatrix(i, 5)) & ", MSALES.OTHERS = " & Val(.TextMatrix(i, 6)) & ", MSALES.TTLSALES = " & Val(.TextMatrix(i, 7)) & _
                                   " WHERE (((MSALES.SDATE)=#" & .TextMatrix(i, 0) & "#));"
                    'mySQL
                    LYRcon.Execute "UPDATE MSALES SET MSALES.CASH = " & Val(.TextMatrix(i, 1)) & ", MSALES.ACCTSALES = " & Val(.TextMatrix(i, 2)) & ", MSALES.SALESRET = " & Val(.TextMatrix(i, 3)) & ", MSALES.SRVCCASH = " & Val(.TextMatrix(i, 4)) & ", MSALES.SRVCACCT = " & Val(.TextMatrix(i, 5)) & ", MSALES.OTHERS = " & Val(.TextMatrix(i, 6)) & ", MSALES.TTLSALES = " & Val(.TextMatrix(i, 7)) & _
                                   " WHERE MSALES.SDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                    
                Else
                    'Insert Into table Monthly Sales
                    'LYRcon.Execute "INSERT INTO MSALES ( SDATE, CASH, ACCTSALES, SALESRET, SRVCCASH, SRVCACCT, OTHERS, TTLSALES ) " & _
                                   "VALUES (#" & .TextMatrix(i, 0) & "#, " & Val(.TextMatrix(i, 1)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & Val(.TextMatrix(i, 5)) & ", " & Val(.TextMatrix(i, 6)) & ", " & Val(.TextMatrix(i, 7)) & ");"
                    'mySQL
                    LYRcon.Execute "INSERT INTO MSALES ( SDATE, CASH, ACCTSALES, SALESRET, SRVCCASH, SRVCACCT, OTHERS, TTLSALES ) " & _
                                   "VALUES ('" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 1)) & ", " & Val(.TextMatrix(i, 2)) & ", " & Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & Val(.TextMatrix(i, 5)) & ", " & Val(.TextMatrix(i, 6)) & ", " & Val(.TextMatrix(i, 7)) & ");"
                    
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

