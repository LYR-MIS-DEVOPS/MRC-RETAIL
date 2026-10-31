VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form InsertPOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "On-Account Customer"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   Icon            =   "InsertPOld.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDP 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer w/ DP"
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   255
      Width           =   1575
   End
   Begin VB.TextBox txtOSBal 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Frame FraTerms 
      Height          =   1755
      Left            =   120
      TabIndex        =   20
      Tag             =   "XYF"
      Top             =   1560
      Width           =   5895
      Begin VB.TextBox txtAmtDue 
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
         Left            =   3240
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   940
         Width           =   1335
      End
      Begin VB.TextBox txtTtlPrice 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   940
         Width           =   1335
      End
      Begin VB.ComboBox cboScheme 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   4215
      End
      Begin VB.TextBox txtStartPay 
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
         Left            =   1560
         TabIndex        =   7
         Tag             =   "XYF"
         Text            =   "08/08/2008"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtTerm 
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
         Left            =   1560
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   560
         Width           =   1335
      End
      Begin VB.TextBox txtGives 
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
         Left            =   3240
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   560
         Width           =   1335
      End
      Begin VB.Label lblPayTerm 
         Caption         =   "MONTHLY"
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
         Left            =   4680
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   970
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "@"
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
         Left            =   2955
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   970
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "TOTAL PRICE"
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
         Left            =   120
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   970
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "PAYMENT START"
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
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   1350
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "PAYMENT SCHEME"
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
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "TERMS [DAYS]"
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
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   590
         Width           =   1695
      End
      Begin VB.Label lblGives 
         Caption         =   "PAYMENTS"
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
         Left            =   4680
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox cboBranch 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   585
      Width           =   4215
   End
   Begin VB.TextBox txtPO 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   840
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Text            =   "Qty. Purchased: 0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   3360
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
      Image           =   "InsertPOld.frx":058A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3360
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
      Image           =   "InsertPOld.frx":0B24
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      Caption         =   "O.S. BALANCE"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   28
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "ADDRESS"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "CUSTOMER"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   615
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "P.O. NUMBER:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   255
      Width           =   1695
   End
   Begin VB.Label lblX 
      Caption         =   "û"
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
      Left            =   2760
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
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
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuSP 
      Caption         =   "SP"
      Visible         =   0   'False
      Begin VB.Menu mnuMark 
         Caption         =   "Mark"
      End
      Begin VB.Menu mnuUnmark 
         Caption         =   "Unmark"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkAll 
         Caption         =   "Mark All"
      End
      Begin VB.Menu mnuUnmarkAll 
         Caption         =   "UnMark All"
      End
   End
End
Attribute VB_Name = "InsertPOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCustomer As New ADODB.Recordset
Dim PrtSrch As Boolean

Function GetOR(MyID As Integer) As String
    Dim rsDP As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE (((PAYMENTS.DOWNPAYMENT)>0) AND ((PAYMENTS.CUSTOMERID)=" & MyID & ")) ORDER BY PAYMENTS.PAYDATE DESC;"
    With rsDP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetOR = IIf(IsNull(!ORNO), "", !ORNO)
        Else
            GetOR = ""
        End If
        .Close
    End With
End Function


Private Sub cboBranch_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
        
    If PrtSrch Then
        PrtSrch = False
        Exit Sub
    End If
    With cboBranch
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsCustomer.MoveFirst
            rsCustomer.Move i
            txtCustomerID.Text = rsCustomer!CUSTOMERID
            txtAddress.Text = IIf(IsNull(rsCustomer!Address), "", rsCustomer!Address)
            txtOSBal.Text = Format$(rsCustomer!OSbal, "#,###,##0.00")
            If chkDP.Value = 0 Then
                txtTtlPrice.Text = SPSorder.txtTotal.Text
            Else
                txtPO.Text = GetOR(txtCustomerID.Text)
                If txtPO.Text <> "" Then Label1(0).Caption = "O.R. NUMBER"
            End If
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PrtSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtAddress.Text = ""
            txtOSBal.Text = ""
        End If
    End With
End Sub

Private Sub cboBranch_Click()
    Call cboBranch_Change
End Sub

Private Sub cboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PrtSrch = True
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboScheme_Change()
    Dim iPS As Integer, TPrice As Currency
    Dim m As Integer, w As Integer, sPay As Date
    
    If cboScheme.Text = "WEEKLY" Then iPS = 7
    If cboScheme.Text = "BI-MONTHLY" Then iPS = 15
    If cboScheme.Text = "MONTHLY" Then iPS = 30
    TPrice = Val(Replace(txtTtlPrice.Text, ",", ""))
    If Val(txtGives.Text) > 0 Then  '#Payments available
        'Compute for term base on scheme
        If iPS = 7 Then
            m = 0
            If Val(txtGives.Text) >= 4 Then
                m = Val(txtGives.Text) \ 4
                w = Val(txtGives.Text) Mod 4
            Else
                w = Val(txtGives.Text)
            End If
            txtTerm.Text = (m * 30) + (w * 7)
        Else
            txtTerm.Text = Val(txtGives.Text) * iPS
        End If
        If TPrice > 0 Then
            txtAmtDue.Text = Format$(TPrice / Val(txtGives.Text), "#,##0.00")
        End If
    ElseIf Val(txtTerm.Text) > 0 Then
        If Val(txtTerm.Text) >= iPS Then
            txtGives.Text = Val(txtTerm.Text) \ iPS
        Else
            txtGives.Text = 1
        End If
    End If
    If iPS = 15 Then
        lblPayTerm.Caption = "15 DAYS"
    Else
        lblPayTerm.Caption = cboScheme.Text
    End If
    If iPS = 7 Then
        If Day(SPSorder.txtDate.Text) < 5 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 7)
        ElseIf Day(SPSorder.txtDate.Text) < 12 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 14)
        ElseIf Day(SPSorder.txtDate.Text) < 19 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 21)
        ElseIf Day(SPSorder.txtDate.Text) < 28 Then
            sPay = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, SPSorder.txtDate.Text)), Month(DateAdd("m", 1, SPSorder.txtDate.Text)), 1))
        Else
            sPay = DateAdd("m", 1, DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 7))
        End If
        txtStartPay.Text = Format$(sPay, "mm/dd/yyyy")
    Else
        txtStartPay.Text = DateAdd("d", iPS, SPSorder.txtDate.Text)
    End If
        
End Sub

Private Sub cboScheme_Click()
    Call cboScheme_Change
End Sub

Private Sub cboScheme_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGives.SetFocus
End Sub

Private Sub chkDP_Click()
    If chkDP.Value = 1 Then
        FraTerms.Enabled = False
        cboScheme.ListIndex = -1
        txtTerm.Text = ""
        txtGives.Text = ""
        txtTtlPrice.Text = ""
        txtAmtDue.Text = ""
        txtStartPay.Text = ""
    Else
        FraTerms.Enabled = False
        Label1(0).Caption = "P.O. NUMBER"
    End If

End Sub

Private Sub cmdCancel_Click()
    Dim Ans As Integer
    
    If SPSorder.txtSPSorderID.Text <> "" Then
        If SPSorder.txtCustID.Text <> "" Then
            Ans = MsgBox("Do you want to remove On-Account Customer?", vbQuestion + vbYesNo, "Remove On Account")
            If Ans = vbYes Then
                With SPSorder
                    .txtRefNo.Text = ""
                    .lblRef.Caption = "Ref. No."
                    .txtPONO = ""
                    .txtCustID.Text = ""
                    .txtCustName.Text = ""
                    .txtPayScheme.Text = ""
                    .txtTerms.Text = ""
                    .txtNPays.Text = ""
                    .txtTotal.Text = ""
                    .txtBillAmt.Text = ""
                    .txtPayStart.Text = ""
                    .chkDP.Value = 0
                End With
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim Ans As Integer
    Dim vMsg As String, vQry As String
    
    vMsg = ""
    If txtCustomerID.Text = "" Then vMsg = "Please select a Customer from list"
    If chkDP.Value = 1 Then
        If txtCustomerID.Text <> "" Then
            vQry = "SELECT DPLEDGER.* From DPLEDGER WHERE (((DPLEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((DPLEDGER.STATUS)='OPEN'));"
            If Not CtrlNumExist(vQry) Then vMsg = "No Down Payment made by " & cboBranch.Text
        End If
    Else
        If txtPO.Text <> "" And Val(txtCustomerID.Text) > 0 Then
            vQry = "SELECT CUSTPO.CUSTOMERID From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtPO.Text & "'));"
            Ans = Val(GetFieldValue(vQry))
            If Ans > 0 Then
                If Ans <> Val(txtCustomerID.Text) Then vMsg = "P.O. No. " & txtPO.Text & " does not match Custmer Name."
            End If
        End If
        If cboScheme.Text = "" Then vMsg = vMsg & vbCrLf & "No payment scheme specified"
        If Val(txtGives.Text) = 0 Then vMsg = vMsg & vbCrLf & "Please specify number of payments"
        If Not IsDate(txtStartPay.Text) Then vMsg = vMsg & vbCrLf & "Invalid date of payment start"
        If Val(txtAmtDue.Text) = 0 Then vMsg = vMsg & vbCrLf & "No Amount due specified"
    End If
    If vMsg <> "" Then
        MsgBox vMsg, vbCritical + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    With SPSorder
        If Val(.txtReqID.Text) = 0 Then
            .txtRefNo.Text = txtPO.Text
            .lblRef.Caption = "P.O. Number"
            .lblCustomer.Caption = "On-Acct. Customer"
            .txtPONO = txtPO.Text
        End If
        If chkDP.Value = 1 Then
            .txtRefNo.Text = txtPO.Text
            .lblRef.Caption = "O.R. Number"
            .lblCustomer.Caption = "Customer w/ DP"
            .txtPONO = txtPO.Text
        End If
        .txtCustID.Text = txtCustomerID.Text
        .txtCustName.Text = cboBranch.Text
        .txtPayScheme.Text = cboScheme.Text
        .txtTerms.Text = txtTerm.Text
        .txtNPays.Text = txtGives.Text
        .txtTotal.Text = txtTtlPrice.Text
        .txtBillAmt.Text = txtAmtDue.Text
        .txtPayStart.Text = txtStartPay.Text
        .chkDP.Value = chkDP.Value
    End With
    Unload Me
End Sub

Private Sub ShowPO()
    Dim rsPO As New ADODB.Recordset
    Dim rsPOdtls As New ADODB.Recordset
    Dim Qry As String, myMsg As String, MyInv As String, myRem As String
    Dim POid As Integer
    
    myMsg = ""
    Qry = "SELECT CUSTPO.* From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtPO.Text & "'));"
    With rsPO
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            POid = !CPOID
            cboBranch.Text = !CUSTOMER
            MyInv = IIf(IsNull(!INVOICENO), "", !INVOICENO)
            myRem = IIf(IsNull(!Remarks), "", !Remarks)
            If MyInv <> "" Then
                myMsg = "P.O. no. " & txtPO.Text & " was already issued with Invoice No. " & MyInv
            ElseIf !Approved = False Then
                myMsg = "P.O. no. " & txtPO.Text & " is not yet approved."
            ElseIf InStr(myRem, "CANCELLED") > 0 Then
                myMsg = "P.O. no. " & txtPO.Text & " was Cancelled!"
            Else
                SPSorder.txtCPOid.Text = POid
                SPSorder.lblRef.Caption = "P.O. Number"
                SPSorder.txtRefNo.Text = txtPO.Text
                SPSorder.txtCustName.Text = !CUSTOMER
                SPSorder.txtCustID.Text = !CUSTOMERID
                SPSorder.txtJOid.Text = IIf(IsNull(!JoborderID), "", !JoborderID)
                SPSorder.lblCustomer = "On-Acct. Customer"
            End If
        End If
        .Close
    End With
    If myMsg <> "" Then
        MsgBox myMsg, vbExclamation + vbOKOnly, "Customer P.O."
        txtPO.Text = ""
        Exit Sub
    End If
    'Display Ordered Parts
    With SPSorder
        .flxOrd.Rows = 1
        .flxOrd.Rows = 2
        Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID " & _
              "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.CPOID)=" & POid & "));"
        rsPOdtls.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsPOdtls.RecordCount > 0 Then
            rsPOdtls.MoveFirst
            .flxOrd.Rows = 1
            Do While Not rsPOdtls.EOF
                .flxOrd.AddItem vbTab & rsPOdtls!QtyOnHand & vbTab & rsPOdtls!Qty & vbTab & IIf(IsNull(rsPOdtls!Unit), "", rsPOdtls!Unit) & vbTab & rsPOdtls!PartNumber & vbTab & rsPOdtls!Description & vbTab & vbTab & Format$(rsPOdtls!Price, "###,##0.00") & vbTab & Format$(rsPOdtls!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(rsPOdtls!Location), "", rsPOdtls!Location) & vbTab & rsPOdtls!SparepartID & vbTab & IIf(IsNull(rsPOdtls!ApplicModel), "", rsPOdtls!ApplicModel)
                rsPOdtls.MoveNext
            Loop
        End If
        rsPOdtls.Close
        Call .ComputeTotal
        txtTtlPrice.Text = .txtTotal.Text
    End With
    Qry = "SELECT ONACCTERMS.TERMS From ONACCTERMS WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & "));"
    txtTerm.Text = GetFieldValue(Qry)
    cboScheme.Text = "MONTHLY"
    'txtGives.Text = 1
    Call txtGives_LostFocus
End Sub

Private Sub Form_Load()
    Dim Qry As String
    
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL " & _
          "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"

    With rsCustomer
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboBranch.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
    cboScheme.AddItem "WEEKLY"
    cboScheme.AddItem "BI-MONTHLY"
    cboScheme.AddItem "MONTHLY"
    With SPSorder
        If .txtCustID.Text <> "" Then
            txtPO.Text = .txtRefNo.Text
            txtCustomerID.Text = .txtCustID.Text
            cboBranch.Text = .txtCustName.Text
            txtOSBal.Text = Format$(.txtOSBal.Text, "#,###,##0.00")
            If .txtPayScheme.Text <> "" Then
                cboScheme.Text = .txtPayScheme.Text
                txtTerm.Text = .txtTerms.Text
                txtGives.Text = .txtNPays.Text
                txtTtlPrice.Text = .txtTotal.Text
                txtAmtDue.Text = Format$(.txtBillAmt.Text, "#,##0.00")
                txtStartPay.Text = .txtPayStart.Text
                If Val(txtTtlPrice.Text) > 0 And Val(txtGives.Text) > 0 And Val(txtAmtDue.Text) = 0 Then Call txtGives_LostFocus
            Else
                Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE (((DPLEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((DPLEDGER.STATUS)='OPEN'));"
                If CtrlNumExist(Qry) Then chkDP.Value = 1
            End If
        End If
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsCustomer.State = adStateOpen Then rsCustomer.Close
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtAmtDue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStartPay.SetFocus
End Sub

Private Sub txtGives_KeyPress(KeyAscii As Integer)
    Dim StrValid As String
    
    StrValid = "1234567890"
    If keysacii > 26 Then
        If InStr(StrValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtAmtDue.SetFocus
End Sub

Private Sub txtGives_LostFocus()
    Dim iPS As Integer, TPrice As Currency, sPay As Date
    
    If cboScheme.Text = "" Or Val(txtGives.Text) = 0 Then Exit Sub
    If cboScheme.Text = "WEEKLY" Then iPS = 7
    If cboScheme.Text = "BI-MONTHLY" Then iPS = 15
    If cboScheme.Text = "MONTHLY" Then iPS = 30
    TPrice = Val(Replace(txtTtlPrice.Text, ",", ""))
        'Compute for term base on scheme
    If iPS = 7 Then
        m = 0
        If Val(txtGives.Text) >= 4 Then
            m = Val(txtGives.Text) \ 4
            w = Val(txtGives.Text) Mod 4
        Else
            w = Val(txtGives.Text)
        End If
        txtTerm.Text = (m * 30) + (w * 7)
    Else
        txtTerm.Text = Val(txtGives.Text) * iPS
    End If
    If TPrice > 0 Then
        txtAmtDue.Text = Format$(TPrice / Val(txtGives.Text), "#,##0.00")
    End If
    If iPS = 15 Then
        lblPayTerm.Caption = "15 DAYS"
    Else
        lblPayTerm.Caption = cboScheme.Text
    End If
    If iPS = 7 Then
        If Day(SPSorder.txtDate.Text) < 5 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 7)
        ElseIf Day(SPSorder.txtDate.Text) < 12 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 14)
        ElseIf Day(SPSorder.txtDate.Text) < 19 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 21)
        ElseIf Day(SPSorder.txtDate.Text) < 28 Then
            sPay = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, SPSorder.txtDate.Text)), Month(DateAdd("m", 1, SPSorder.txtDate.Text)), 1))
        Else
            sPay = DateAdd("m", 1, DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 7))
        End If
        txtStartPay.Text = Format$(sPay, "mm/dd/yyyy")
    Else
        txtStartPay.Text = DateAdd("d", iPS, SPSorder.txtDate.Text)
    End If
End Sub

Private Sub txtOSBal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboScheme.SetFocus
End Sub

Private Sub txtPO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPO_LostFocus()
    Dim Qry As String
    
    If txtPO.Text = "" Then Exit Sub
    Qry = "SELECT CUSTPO.*, CUSTPO.CPONO From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtPO.Text & "'));"
    If CtrlNumExist(Qry) Then Call ShowPO
End Sub

Private Sub txtStartPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    Dim StrValid As String
    
    StrValid = "1234567890"
    If keysacii > 26 Then
        If InStr(StrValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub txtTerm_LostFocus()
    Dim iPS As Integer, TPrice As Currency, sPay As Date
    
    If cboScheme.Text = "" Then Exit Sub
    If cboScheme.Text = "WEEKLY" Then iPS = 7
    If cboScheme.Text = "BI-MONTHLY" Then iPS = 15
    If cboScheme.Text = "MONTHLY" Then iPS = 30
    TPrice = PureNum(txtTtlPrice.Text)
    If Val(txtTerm.Text) >= iPS Then
        txtGives.Text = Val(txtTerm.Text) \ iPS
    Else
        txtGives.Text = 1
    End If
    If iPS = 15 Then
        lblPayTerm.Caption = "15 DAYS"
    Else
        lblPayTerm.Caption = cboScheme.Text
    End If
    If iPS = 7 Then
        If Day(SPSorder.txtDate.Text) < 5 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 7)
        ElseIf Day(SPSorder.txtDate.Text) < 12 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 14)
        ElseIf Day(SPSorder.txtDate.Text) < 19 Then
            sPay = DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 21)
        ElseIf Day(SPSorder.txtDate.Text) < 28 Then
            sPay = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, SPSorder.txtDate.Text)), Month(DateAdd("m", 1, SPSorder.txtDate.Text)), 1))
        Else
            sPay = DateAdd("m", 1, DateSerial(Year(SPSorder.txtDate.Text), Month(SPSorder.txtDate.Text), 7))
        End If
        txtStartPay.Text = Format$(sPay, "mm/dd/yyyy")
    Else
        txtStartPay.Text = DateAdd("d", iPS, SPSorder.txtDate.Text)
    End If

End Sub

