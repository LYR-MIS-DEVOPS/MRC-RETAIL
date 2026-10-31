VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TRAImport 
   Caption         =   "Import Trust Receipt Agreement"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   5280
   End
   Begin MSFlexGridLib.MSFlexGrid flxCust 
      Height          =   975
      Left            =   9720
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   6
      FormatString    =   "      |<CUSTID    |<CTYPE        |<CUSTNAME   |<ADDRESS          |<TERM       "
   End
   Begin MSFlexGridLib.MSFlexGrid FlxPay 
      Height          =   735
      Left            =   10440
      TabIndex        =   11
      Top             =   5040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   10
      FormatString    =   "    |<PayDate|<InvoiceNO|<SpareParts|<CustomerName|<CustomerID|<Remarks|<EmpId|<UpdatedBy|<LastUpdate"
   End
   Begin MSFlexGridLib.MSFlexGrid flxAcct 
      Height          =   975
      Left            =   10440
      TabIndex        =   10
      Tag             =   "XYF"
      Top             =   3960
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   18
      FormatString    =   $"TRAimport.frx":0000
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   5160
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxTRA 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "     |^TRA DATE  |^TRA NUMBER |^TRA TYPE |^TRA AMOUNT |<CUSTOMER NAME  |<ADDRESS          |<CUST. TYPE "
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Tag             =   "xyf"
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Image           =   "TRAimport.frx":00B4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
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
      Image           =   "TRAimport.frx":064E
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxInv 
      Height          =   1095
      Left            =   10440
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   18
      FormatString    =   $"TRAimport.frx":0BE8
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPITR 
      Height          =   1215
      Left            =   10440
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   15
      FormatString    =   $"TRAimport.frx":0C91
   End
   Begin MSFlexGridLib.MSFlexGrid FlxSPorder 
      Height          =   1335
      Left            =   10440
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   360
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   17
      FormatString    =   $"TRAimport.frx":0D19
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
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   600
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   10410
      m_MDI_Height    =   6435
   End
   Begin VB.Label lblBranch 
      Alignment       =   1  'Right Justify
      Caption         =   "MOBILE VAN"
      Height          =   255
      Left            =   7920
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "DATE COVERED:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   240
      Width           =   4575
   End
   Begin VB.Menu mnuFlxTra 
      Caption         =   "TRA menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMark 
         Caption         =   "Mark"
      End
      Begin VB.Menu mnuUnmark 
         Caption         =   "Unmark"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkAll 
         Caption         =   "Mark All"
      End
      Begin VB.Menu mnuUnmarkAll 
         Caption         =   "Unmark All"
      End
   End
End
Attribute VB_Name = "TRAImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VANcon As New ADODB.Connection

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsOrder As New ADODB.Recordset
    Dim rsInv As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim rsAcct As New ADODB.Recordset
    Dim i As Integer
    Dim Qry As String
        
    
    FlxSPorder.Rows = 1
    Pbar.Value = 0
    Pbar.Max = flxTRA.Rows * 4
    For i = 1 To flxTRA.Rows - 1
        'Display table SPsalesOrder in FlxSPOrder
        If flxTRA.TextMatrix(i, 0) <> "" Then
            With rsOrder
                Qry = "SELECT spsalesorder.SPSORDERID, spsalesorder.ORDERNO, spsalesorder.ORDERDATE, spsalesorder.FRONTLINER, spsalesorder.REFERENCE, spsalesorder.REMARKS, spsalesorder.GROSSPRICE, spsalesorder.PCTDISC, spsalesorder.LESSDISCPRICE, spsalesorder.TOTALPRICE, " & _
                      "spsalesorder.CUSTOMERID, spsalesorder.CTYPE, spsalesorder.CUSTOMER, spsalesorder.TERMS, spsalesorder.ADDRESS, spsalesorder.WORKSTATION From spsalesorder WHERE spsalesorder.SPSORDERID = " & Val(flxTRA.TextMatrix(i, 10)) & ";"
                .Open Qry, VANcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        FlxSPorder.AddItem vbTab & !SPSORDERID & vbTab & !ORDERNO & vbTab & Format(!ORDERDATE, "yyyy-mm-dd") & vbTab & !Frontliner & vbTab & !REFERENCE & vbTab & !Remarks & vbTab & !GROSSPRICE & vbTab & !PctDisc & vbTab & !LESSDISCPRICE & vbTab & !TOTALPRICE & vbTab & !CUSTOMERID & vbTab & !Ctype & vbTab & !CUSTOMER & vbTab & !Terms & vbTab & !Address & vbTab & !WORKSTATION
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        End If
        Pbar.Value = Pbar.Value + 1
    Next i
    
    'Update unrecorded customer
    flxCust.Rows = 1
    For i = 1 To FlxSPorder.Rows - 1
        If Val(FlxSPorder.TextMatrix(i, 11)) > 0 Then
            Qry = "SELECT customer.CUSTOMERID From customer Where customer.CUSTOMERNAME = '" & FlxSPorder.TextMatrix(i, 13) & "';"
            If GetFieldValue(Qry) = "" Then
                flxCust.AddItem vbTab & FlxSPorder.TextMatrix(i, 11) & vbTab & FlxSPorder.TextMatrix(i, 12) & vbTab & FlxSPorder.TextMatrix(i, 13) & vbTab & FlxSPorder.TextMatrix(i, 15) & vbTab & FlxSPorder.TextMatrix(i, 14)
            End If
        End If
    Next i
    
    flxInv.Rows = 1
    For i = 1 To flxTRA.Rows - 1
        If flxTRA.TextMatrix(i, 0) <> "" Then
            With rsInv
                Qry = "SELECT SPINVNO, Transno, InvDate, ChargeTo, InvType, Address, Amount, Frontliner, EmpId, SPSORDERID,  Ctype, ReqID, Remarks, CUSTOMERID, UpdatedBy, LastUpdate, PrintCopy " & _
                    "From spinvoice WHERE spinvoice.SPSORDERID = " & Val(flxTRA.TextMatrix(i, 10)) & ";"
                .Open Qry, VANcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        flxInv.AddItem vbTab & !SPINVNO & vbTab & !TransNo & vbTab & Format(!InvDate, "yyyy-mm-dd") & vbTab & !ChargeTo & vbTab & !InvType & vbTab & !Address & vbTab & !Amount & vbTab & !Frontliner & vbTab & !EMPID & vbTab & !SPSORDERID & vbTab & !Ctype & vbTab & !ReqID & vbTab & !Remarks & vbTab & !CUSTOMERID & vbTab & !UPDATEDBY & vbTab & !LASTUPDATE & vbTab & !PRINTCOPY
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        End If
        Pbar.Value = Pbar.Value + 1
    Next i
 
    flxSPITR.Rows = 1
    For i = 1 To flxTRA.Rows - 1
        If flxTRA.TextMatrix(i, 0) <> "" Then
            With rsITR
                Qry = "SELECT Reference, TRANSTYPE, Remarks, TransDate, SparepartID, `OUT`, Price, Discount, UnitM, Cost, Source, UpdatedBy, LastUpdate, `Order` " & _
                    "FROM spitr Where spitr.REFERENCE = '" & flxTRA.TextMatrix(i, 2) & "';"
                .Open Qry, VANcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        flxSPITR.AddItem vbTab & !REFERENCE & vbTab & !TRANSTYPE & vbTab & !Remarks & vbTab & !TRANSDATE & vbTab & !SparepartID & vbTab & !OUT & vbTab & !Price & vbTab & !DISCOUNT & vbTab & !UNITM & vbTab & !Cost & vbTab & !Source & vbTab & !UPDATEDBY & vbTab & !LASTUPDATE & vbTab & !Order
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        End If
        Pbar.Value = Pbar.Value + 1
    Next i
                
    
    flxAcct.Rows = 1
    FlxPay.Rows = 1
    For i = 1 To flxTRA.Rows - 1
        If flxTRA.TextMatrix(i, 0) <> "" And flxTRA.TextMatrix(i, 3) = "CREDIT" Then
            With rsAcct
                Qry = "SELECT SPINVNO, Terms, PayScheme, Npays, PayStart, BillingAmt, NxtBill, AmtDue, Status, TtlAmtPaid, BALANCE, SPSORDERID, CUSTOMERID, UpdatedBy, LastUpdate, AcctAmount, TransDate " & _
                    "FROM onaccterms Where onaccterms.SPSORDERID = " & Val(flxTRA.TextMatrix(i, 10)) & " AND onaccterms.SPINVNO = '" & flxTRA.TextMatrix(i, 2) & "';"
                .Open Qry, VANcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        flxAcct.AddItem vbTab & !SPINVNO & vbTab & !Terms & vbTab & !PayScheme & vbTab & !Npays & vbTab & !PayStart & vbTab & !BILLINGAMT & vbTab & !NxtBill & vbTab & !AmtDue & vbTab & !STATUS & vbTab & !TTLAMTPAID & vbTab & !Balance & vbTab & !SPSORDERID & vbTab & !CUSTOMERID & vbTab & !UPDATEDBY & vbTab & !LASTUPDATE & vbTab & !AcctAmount & vbTab & !TRANSDATE
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        ElseIf flxTRA.TextMatrix(i, 0) <> "" And flxTRA.TextMatrix(i, 3) = "CASH" Then
            With rsAcct
                Qry = "SELECT PayDate, InvoiceNO, SpareParts, CustomerName, CustomerID, Remarks, EmpId, UpdatedBy, LastUpdate " & _
                      "FROM payments Where payments.INVOICENO = '" & flxTRA.TextMatrix(i, 2) & "';"
                .Open Qry, VANcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        FlxPay.AddItem vbTab & !PayDate & vbTab & !InvoiceNO & vbTab & !SPAREPARTS & vbTab & !CustomerName & vbTab & !CUSTOMERID & vbTab & !Remarks & vbTab & !EMPID & vbTab & !UPDATEDBY & vbTab & !LASTUPDATE
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        End If
        Pbar.Value = Pbar.Value + 1
    Next i
    
    'Saving Process
    Pbar.Value = 0
    Pbar.Max = FlxSPorder.Rows + flxSPITR.Rows + flxInv.Rows + flxAcct.Rows + FlxPay.Rows
    
    With flxCust
        For i = 1 To .Rows - 1
            LYRcon.Execute "INSERT INTO customer(CUSTOMERID, CTYPE, CUSTOMERNAME, ADDRESS, TERMS, APPLYSERVICE, UPDATEDBY, LASTUPDATE) " & _
            "VALUES(" & Val(.TextMatrix(i, 1)) & ", '" & .TextMatrix(i, 2) & "', '" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 4) & "', " & _
            Val(.TextMatrix(i, 5)) & ", 0, " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        Next i
    End With
        
    With FlxSPorder
        For i = 1 To .Rows - 1
            LYRcon.Execute "Insert Into spsalesorder(SPSORDERID, ORDERNO, ORDERDATE, FRONTLINER, REFERENCE, REMARKS, GROSSPRICE, PCTDISC, LESSDISCPRICE, TOTALPRICE, CUSTOMERID, CTYPE, CUSTOMER, TERMS, ADDRESS, WORKSTATION) " & _
            "VALUES(" & Val(.TextMatrix(i, 1)) & ", " & Val(.TextMatrix(i, 2)) & ", '" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 4) & "', '" & .TextMatrix(i, 5) & "', '" & .TextMatrix(i, 6) & "', " & Val(.TextMatrix(i, 7)) & _
            ", " & Val(.TextMatrix(i, 8)) & ", " & Val(.TextMatrix(i, 9)) & ", " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 11)) & ", '" & .TextMatrix(i, 12) & "', '" & .TextMatrix(i, 13) & "', " & Val(.TextMatrix(i, 14)) & ", '" & .TextMatrix(i, 15) & "', '" & .TextMatrix(i, 16) & "')"
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    With flxInv
        For i = 1 To .Rows - 1
            LYRcon.Execute "INSERT INTO spinvoice(SPINVNO, Transno, InvDate, ChargeTo, InvType, Address, Amount, Frontliner, EmpId, SPSORDERID, Ctype, ReqID, Remarks, CUSTOMERID, UpdatedBy, LastUpdate, PrintCopy) " & _
            "VALUES('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 4) & "', '" & .TextMatrix(i, 5) & "', '" & .TextMatrix(i, 6) & "', " & Val(.TextMatrix(i, 7)) & _
            ", '" & .TextMatrix(i, 8) & "', " & Val(.TextMatrix(i, 9)) & ", " & Val(.TextMatrix(i, 10)) & ", '" & .TextMatrix(i, 11) & "', " & Val(.TextMatrix(i, 12)) & ", '" & .TextMatrix(i, 13) & "', " & Val(.TextMatrix(i, 14)) & ", " & Val(.TextMatrix(i, 15)) & ", '" & Format(DateValue(.TextMatrix(i, 16)), "yyyy-mm-dd hh:mm:ss") & "', " & Val(.TextMatrix(i, 17)) & ")"
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    With flxSPITR
        For i = 1 To .Rows - 1
            LYRcon.Execute "INSERT INTO spitr(Reference, TRANSTYPE, Remarks, TransDate, SparepartID, `OUT`, Price, DISCOUNT, UNITM, COST, Source, UpdatedBy, LastUpdate, `Order`) " & _
            "VALUES('" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '" & .TextMatrix(i, 3) & "', '" & Format(DateValue(.TextMatrix(i, 4)), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 5)) & ", " & Val(.TextMatrix(i, 6)) & ", " & Val(.TextMatrix(i, 7)) & _
            ", " & Val(.TextMatrix(i, 8)) & ", '" & .TextMatrix(i, 9) & "', " & Val(.TextMatrix(i, 10)) & ", '" & .TextMatrix(i, 11) & "', " & Val(.TextMatrix(i, 12)) & ", '" & Format(DateValue(.TextMatrix(i, 13)), "yyyy-mm-dd hh:mm:ss") & "', " & Val(.TextMatrix(i, 14)) & ")"
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    With flxAcct
        For i = 1 To .Rows - 1
            LYRcon.Execute "INSERT INTO onaccterms(SPINVNO, Terms, PayScheme, Npays, PayStart, BillingAmt, NxtBill, AmtDue, `Status`, TtlAmtPaid, BALANCE, SPSORDERID, CUSTOMERID, UpdatedBy, LastUpdate, AcctAmount, TransDate) " & _
            "VALUES('" & .TextMatrix(i, 1) & "', " & Val(.TextMatrix(i, 2)) & ", '" & .TextMatrix(i, 3) & "', " & Val(.TextMatrix(i, 4)) & ", '" & Format(DateValue(.TextMatrix(i, 5)), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 6)) & ", '" & Format(DateValue(.TextMatrix(i, 7)), "yyyy-mm-dd") & _
            "', " & Val(.TextMatrix(i, 8)) & ", '" & .TextMatrix(i, 9) & "', " & Val(.TextMatrix(i, 10)) & ", " & Val(.TextMatrix(i, 11)) & ", " & Val(.TextMatrix(i, 12)) & ", " & Val(.TextMatrix(i, 13)) & ", " & Val(.TextMatrix(i, 14)) & ", '" & Format(DateValue(.TextMatrix(i, 15)), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 16)) & ", '" & Format(DateValue(.TextMatrix(i, 17)), "yyyy-mm-dd") & "')"
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    With FlxPay
        For i = 1 To .Rows - 1
            If IsDate(.TextMatrix(i, 1)) Then
                LYRcon.Execute "INSERT INTO payments(PayDate, InvoiceNO, SpareParts, CustomerName, CustomerID, Remarks, EmpId, UpdatedBy, LastUpdate) " & _
                "VALUES('" & Format(DateValue(.TextMatrix(i, 1)), "yyyy-mm-dd") & "', '" & .TextMatrix(i, 2) & "', " & Val(.TextMatrix(i, 3)) & ", '" & .TextMatrix(i, 4) & "', " & Val(.TextMatrix(i, 5)) & _
                ", '" & .TextMatrix(i, 6) & "', " & Val(.TextMatrix(i, 7)) & ", " & Val(.TextMatrix(i, 8)) & ", '" & Format(DateValue(.TextMatrix(i, 9)), "yyyy-mm-dd hh:mm:ss") & "')"
            End If
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
        
    
    MsgBox "TRA successfully transferred from VAN Mobile to Office!", vbInformation + vbOKOnly, "Import TRA"
    Call FrmSummary.RefreshRecord(2)
    Unload Me
'LocalErr:

End Sub

Private Sub flxTRA_DblClick()
    If flxTRA.TextMatrix(flxTRA.Row, 0) = "" And IsDate(flxTRA.TextMatrix(flxTRA.Row, 1)) Then
        Call mnuMark_Click
    Else
        Call mnuUnmark_Click
    End If
End Sub

Private Sub flxTRA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxTRA.TextMatrix(flxTRA.Row, 0) = "" And IsDate(flxTRA.TextMatrix(flxTRA.Row, 1)) Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
    End If
End Sub

Private Sub flxTRA_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If flxTRA.TextMatrix(flxTRA.Row, 0) = "" Then
            mnuMark.Enabled = True
            mnuUnmark.Enabled = False
        Else
            mnuMark.Enabled = False
            mnuUnmark.Enabled = True
        End If
        Call PopupMenu(mnuFlxTra)
    End If
End Sub

Private Sub Form_Load()

    VANcon.CursorLocation = adUseClient
    VANcon.ConnectionTimeout = 60
    If UCase$(SysBranch) = "LYR MOTORPARTS CORP. - VAN-2 [LHC403]" Then
        VANcon.Open "dsn=van1mobile"
        lblBranch.Caption = "VAN-2 MOBILE [LHC403]"
    ElseIf UCase$(SysBranch) = "LYR MOTORPARTS CORP. - VAN-56 [AAH1233]" Then
        VANcon.Open "dsn=van2mobile"
        lblBranch.Caption = "VAN-56 MOBILE [AAH1233]"
    ElseIf UCase$(SysBranch) = "LYR MOTORPARTS CORP. - VAN-55 [AAH1250]" Then
        VANcon.Open "dsn=van3mobile"
        lblBranch.Caption = "VAN-55 MOBILE [AAH1250]"
    End If
    lblBranch.Caption = SysBranch
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If VANcon.State = adStateOpen Then VANcon.Close
End Sub

Private Sub mnuMark_Click()
    With flxTRA
        .Col = 0
        .CellFontName = "WINGDINGS"
        .CellFontSize = 12
        .Text = lblChk.Caption
        .CellAlignment = 4
    End With
End Sub

Private Sub mnuMarkAll_Click()
    Dim i As Integer
    
    With flxTRA
        For i = 1 To .Rows - 1
            .Row = i
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .TextMatrix(i, 0) = lblChk.Caption
        Next i
    End With
End Sub

Private Sub mnuUnmark_Click()
    flxTRA.TextMatrix(flxTRA.Row, 0) = ""
End Sub

Private Sub mnuUnmarkAll_Click()
    Dim i As Integer
    
    For i = 1 To flxTRA.Rows - 1
        flxTRA.TextMatrix(i, 0) = ""
    Next i
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxTRA
        .FormatString = "  |^TRA. DATE|^TRA. NO. |^TRA TYPE|>TRA AMOUNT |<NAME OF CUSTOMER |<ADDRESS |^CUSTOMER TYPE|<REMARKS |<AGENT |^SPID "
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.08
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.09
        .ColWidth(4) = (.Width - 300) * 0.09
        .ColWidth(5) = (.Width - 300) * 0.2
        .ColWidth(6) = (.Width - 300) * 0.2
        .ColWidth(7) = (.Width - 300) * 0.09
        .ColWidth(8) = (.Width - 300) * 0.25
        .ColWidth(9) = (.Width - 300) * 0.2
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        
    End With
End Sub

Private Sub Timer1_Timer()
    Dim rsTRA As New ADODB.Recordset
    Dim Qry As String, i As Integer, j As Integer
    
    
    If Val(FrmSummary.cmdNext.Tag) = 1 Then
        Label1.Caption = "DATE COVERED: " & Format(FrDate, "MMMM DD, YYYY")
        Qry = "SELECT spinvoice.INVDATE, spinvoice.TRANSNO, spinvoice.INVTYPE, spinvoice.AMOUNT, spinvoice.CHARGETO, spinvoice.ADDRESS, spinvoice.CTYPE, spinvoice.REMARKS, spinvoice.FRONTLINER, spinvoice.SPSORDERID " & _
              "From spinvoice Where spinvoice.TRANSNO LIKE 'TRA#%' AND spinvoice.INVDATE = '" & Format(FrDate, "yyyy-mm-dd") & "' AND spinvoice.PRINTCOPY>0 Order By spinvoice.InvDate Desc"
    Else
        Label1.Caption = "DATE COVERED: " & Format(FrDate, "MM/DD/YYYY") & " TO " & Format(ToDate, "MM/DD/YYYY")
        Qry = "SELECT spinvoice.INVDATE, spinvoice.TRANSNO, spinvoice.INVTYPE, spinvoice.AMOUNT, spinvoice.CHARGETO, spinvoice.ADDRESS, spinvoice.CTYPE, spinvoice.REMARKS, spinvoice.FRONTLINER, spinvoice.SPSORDERID " & _
              "From spinvoice Where spinvoice.TRANSNO LIKE 'TRA#%' AND spinvoice.INVDATE >= '" & Format(FrDate, "yyyy-mm-dd") & "' AND spinvoice.INVDATE <= '" & Format(ToDate, "yyyy-mm-dd") & "' AND spinvoice.PRINTCOPY>0 Order By spinvoice.InvDate Desc"
    End If
    With rsTRA
        .Open Qry, VANcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxTRA.Rows = 1
            .MoveFirst
            Do While Not .EOF
                flxTRA.AddItem vbTab & Format(!InvDate, "mm/dd/yyyy") & vbTab & !TransNo & vbTab & !InvType & vbTab & Format(!Amount, "#,##0.00") & vbTab & !ChargeTo & vbTab & !Address & vbTab & !Ctype & vbTab & !Remarks & vbTab & !Frontliner & vbTab & !SPSORDERID '& vbTab & !PrintCopy
                .MoveNext
            Loop
        End If
        .Close
    End With
    With FrmSummary.flxSum  'Check and remove if TRA already Exist
        For i = flxTRA.Rows - 1 To 1 Step -1
            For j = 1 To .Rows - 1
                If flxTRA.TextMatrix(i, 2) = .TextMatrix(j, 2) Then
                    If flxTRA.Rows = 2 And i = 1 Then
                        flxTRA.Rows = 1
                        flxTRA.Rows = 2
                        Exit For
                    Else
                        flxTRA.RemoveItem i
                        Exit For
                    End If
                End If
            Next j
        Next i
    End With
    
    Timer1.Enabled = False
End Sub
