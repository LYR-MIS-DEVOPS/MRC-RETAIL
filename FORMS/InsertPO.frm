VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form InsertPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "On-Account Customer"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "InsertPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPointsType 
      Height          =   315
      ItemData        =   "InsertPO.frx":058A
      Left            =   1800
      List            =   "InsertPO.frx":058C
      TabIndex        =   35
      Top             =   1920
      Width           =   4095
   End
   Begin VB.OptionButton OptCash 
      Caption         =   "CASH"
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.OptionButton OptOnAcct 
      Caption         =   "ON-ACCOUNT"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2280
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox txtCusType 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CheckBox chkDP 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer w/ DP"
      Height          =   255
      Left            =   4320
      TabIndex        =   32
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtOSBal 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1236
      Width           =   4095
   End
   Begin VB.Frame FraTerms 
      Height          =   1755
      Left            =   120
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   2520
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
         TabIndex        =   11
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
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   940
         Width           =   1335
      End
      Begin VB.ComboBox cboScheme 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
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
         TabIndex        =   12
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   914
      Width           =   4095
   End
   Begin VB.ComboBox cboBranch 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   562
      Width           =   4095
   End
   Begin VB.TextBox txtPO 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   840
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   600
      TabIndex        =   16
      Text            =   "Qty. Purchased: 0"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   4320
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
      Image           =   "InsertPO.frx":058E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   4320
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
      Image           =   "InsertPO.frx":0B28
      cBack           =   -2147483633
   End
   Begin MSComCtl2.UpDown UpDRef 
      Height          =   315
      Left            =   1560
      TabIndex        =   36
      Top             =   220
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Max             =   1
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "POINTS TYPE:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   34
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "CUSTOMER TYPE:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   33
      Top             =   1575
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "O.S. BALANCE"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   31
      Top             =   1251
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "ADDRESS"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   929
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "CUSTOMER"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   592
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "P.O. NUMBER:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   19
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
      Left            =   2640
      TabIndex        =   17
      Top             =   5040
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
      Left            =   120
      TabIndex        =   15
      Top             =   5160
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
Attribute VB_Name = "InsertPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCustomer As New ADODB.Recordset
Dim PrtSrch As Boolean
Public isKeyAcc As Boolean
Dim address As String


Function GetOR(MyID As Integer) As String
    Dim rsDP As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE (((PAYMENTS.DOWNPAYMENT)>0) AND ((PAYMENTS.CUSTOMERID)=" & MyId & ")) ORDER BY PAYMENTS.PAYDATE DESC;"
    'mySQL
    Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE PAYMENTS.DOWNPAYMENT>0 AND PAYMENTS.CUSTOMERID=" & MyID & " ORDER BY PAYMENTS.PAYDATE DESC;"
    
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
            txtAddress.Text = IIf(IsNull(rsCustomer!address), "", rsCustomer!address)
            txtOSBal.Text = Format$(rsCustomer!OSbal, "#,###,##0.00")
            txtCusType.Text = IIf(IsNull(rsCustomer!Ctype), "", rsCustomer!Ctype)
            
            'If InStr(txtCusType.Text, "WHOLESALE") > 0 Then
            
            If InStr(txtCusType.Text, "WHOLESALE") > 0 Or InStr(txtCusType.Text, "LGC MARKETING") > 0 Then
                OptOnAcct.Visible = True
                OptCash.Visible = True
            Else
                OptOnAcct.Visible = False
                OptCash.Visible = False
            End If
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
            txtCusType.Text = ""
            OptOnAcct.Visible = False
            OptCash.Visible = False
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
        FraTerms.Enabled = True
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
                    .txtPOno = ""
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
    Dim rsDisc As New ADODB.Recordset
    Dim Ans As Integer, i As Integer
    Dim vMsg As String, vQry As String
    
    vMsg = ""
    If txtCustomerID.Text = "" Then vMsg = "Please select a Customer from list"
    If chkDP.Value = 1 Then
        If txtCustomerID.Text <> "" Then
            'vQry = "SELECT DPLEDGER.* From DPLEDGER WHERE (((DPLEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((DPLEDGER.STATUS)='OPEN'));"
            'mysql
            vQry = "SELECT DPLEDGER.* From DPLEDGER WHERE DPLEDGER.CUSTOMERID=" & txtCustomerID.Text & " AND DPLEDGER.STATUS='OPEN';"
            
            If Not CtrlNumExist(vQry) Then vMsg = "No Down Payment made by " & cboBranch.Text
        End If
    Else
        If txtPO.Text <> "" And Val(txtCustomerID.Text) > 0 Then
            'vQry = "SELECT CUSTPO.CUSTOMERID From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtPO.Text & "'));"
            'mySQL
            vQry = "SELECT CUSTPO.CUSTOMERID From CUSTPO WHERE CUSTPO.CPONO='" & txtPO.Text & "';"
            Ans = Val(GetFieldValue(vQry))
            If Ans > 0 Then
                If Ans <> Val(txtCustomerID.Text) Then vMsg = "P.O. No. " & txtPO.Text & " does not match Custmer Name."
            End If
        End If
        If Not (txtCusType.Text = "WHOLESALE" And OptCash.Value = True) Then
            If cboScheme.Text = "" Then vMsg = vMsg & vbCrLf & "No payment scheme specified"
            If Val(txtGives.Text) = 0 Then vMsg = vMsg & vbCrLf & "Please specify number of payments"
            If Not IsDate(txtStartPay.Text) Then vMsg = vMsg & vbCrLf & "Invalid date of payment start"
            If Val(txtAmtDue.Text) = 0 Then vMsg = vMsg & vbCrLf & "No Amount due specified"
        End If
    End If
    If vMsg <> "" Then
        MsgBox vMsg, vbCritical + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    With SPSorder
    
        If Val(.txtReqID.Text) = 0 Then
            .txtRefNo.Text = txtPO.Text
            If UpDRef.Value > 0 Then
                .lblRef.Caption = "CARD Number"
            Else
                .lblRef.Caption = "P.O. Number"
            End If
            .lblCustomer.Caption = "On-Acct. Customer"
            .txtPOno = txtPO.Text
        End If
        If chkDP.Value = 1 Then
            .txtRefNo.Text = txtPO.Text
            .lblRef.Caption = "O.R. Number"
            .lblCustomer.Caption = "Customer w/ DP"
            .txtPOno = txtPO.Text
            .chkDP.Value = 1
        Else
            .chkDP.Value = 0
        End If
        .txtCustID.Text = txtCustomerID.Text
        .txtCustName.Text = cboBranch.Text
        .cboAddress.Text = GetFieldValue("SELECT CONCAT(customer.CITYTOWN, ', ', customer.PROVINCE) AS ADDRESS From customer Where customer.CUSTOMERID = " & Val(.txtCustID.Text) & ";")
        .txtPayScheme.Text = cboScheme.Text
        .txtTerms.Text = txtTerm.Text
        .txtNPays.Text = txtGives.Text
        .txtTotal.Text = txtTtlPrice.Text
        .txtBillAmt.Text = txtAmtDue.Text
        .txtPayStart.Text = txtStartPay.Text
        .chkDP.Value = chkDP.Value
        .txtCusType = txtCusType.Text
        .myCardPointsType = Trim(cboPointsType.Text) & "/EARN"
        If (UCase$(txtCusType.Text) = "WHOLESALE" Or UCase$(txtCusType.Text) = "LGC MARKETING") And chkDP.Value = 0 Then
            If UCase$(txtCusType.Text) = "WHOLESALE" Then .lblCustomer.Caption = "Wholesale Cust."
            .lblTerms.Caption = "Terms: " & txtTerm.Text & " Days @ " & txtAmtDue.Text & " " & cboScheme.Text
            'Qry = "SELECT SETTINGS.* From SETTINGS " & _
                  "WHERE (((SETTINGS.DESCRIPTION)='Wholesale Discount')) OR (((SETTINGS.DESCRIPTION)='SP Disclusion')) ORDER BY SETTINGS.DESCRIPTION;"
            'mySQL
            Qry = "SELECT SETTINGS.* From SETTINGS " & _
                  "WHERE SETTINGS.DESCRIPTION='Wholesale Discount' OR SETTINGS.DESCRIPTION='SP Disclusion' ORDER BY SETTINGS.DESCRIPTION;"
            
            rsDisc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsDisc.RecordCount > 0 Then
                rsDisc.MoveFirst
                .flxWDisc.Rows = 1
                Do While Not rsDisc.EOF
                    If rsDisc!Description = "Wholesale Discount" Then
                        .flxWDisc.AddItem rsDisc!Remarks & vbTab & rsDisc!Value
                    Else
                        .flxSPec.AddItem rsDisc!Remarks & vbTab & rsDisc!Value
                    End If
                    rsDisc.MoveNext
                Loop
            End If
            rsDisc.Close
            .cmdRemoveWS.Visible = True
            If OptCash.Value = True Then
                .OptCash.Value = True
                .OptOnAcct.Value = False
                .txtPayScheme.Text = ""
                .txtTerms.Text = ""
                .txtNPays.Text = ""
                .txtBillAmt.Text = ""
                .txtPayStart.Text = ""
                .lblTerms.Caption = ""
            Else
                .OptCash.Value = False
                .OptOnAcct.Value = True
                .lblTerms.Caption = "Terms: " & txtTerm.Text & " Days @ " & txtAmtDue.Text & " " & cboScheme.Text
            End If
            If .flxOrd.Rows = 2 And .flxOrd.TextMatrix(1, 10) = "" Then
                Exit Sub
            Else
                If UserID = Val(GetFieldValue("select remarks From settings where description='BM Discount' and value='8/01/2016'")) Then
                    Ans = MsgBox("Do you want the existing orders to avail the wholesale discount?", vbQuestion + vbYesNo, "Wholesale")
                    If Ans = vbNo Then
                        Unload Me
                        Exit Sub
                    End If
                    Call .ComputeWholesaleDiscount
                    txtTtlPrice.Text = .txtTotal.Text
                    Call txtGives_LostFocus
                    .txtBillAmt.Text = txtAmtDue.Text
                    .lblTerms.Caption = "Terms: " & txtTerm.Text & " Days @ " & txtAmtDue.Text & " " & cboScheme.Text
                End If
            End If
        Else
            .flxWDisc.Rows = 0
            .flxSPec.Rows = 0
'            For i = 1 To .flxOrd.Rows - 1
'                If .flxOrd.TextMatrix(i, 6) = "" Then
'                    .flxOrd.TextMatrix(i, 7) = Format$(Val(.flxOrd.TextMatrix(i, 13)), "#,##0.00") 'Orig SRP
'                End If
'            Next i
            Call .ComputeTotal
        End If
    End With
    
     If isKeyAcc Then
        If txtPO.Text = "" Then
            SPSorder.txtRefNo.Text = "KEY ACCOUNTS DISCOUNT"
            SPSorder.cboAddress.Text = txtAddress.Text
            SPSorder.lblTerms.Caption = "Terms Applied"
        Else
            SPSorder.txtRefNo.Text = txtPO.Text
            SPSorder.cboAddress.Text = txtAddress.Text
            SPSorder.lblTerms.Caption = "Terms Applied"
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Qry As String
    
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL, CUSTOMER.CTYPE " & _
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
    cboPointsType.AddItem ""
    cboPointsType.AddItem "CUSTOMER POINTS"
    cboPointsType.AddItem "CLUB POINTS"
    cboPointsType.Enabled = False
    UpDRef.Value = 0
    
    With SPSorder
        If .txtCustID.Text <> "" Then
            txtPO.Text = .txtRefNo.Text
            txtCustomerID.Text = .txtCustID.Text
            cboBranch.Text = .txtCustName.Text
            If isKeyAcc = False Then
            txtOSBal.Text = Format$(.txtOSBal.Text, "#,###,##0.00")
            Else
            cboBranch.Enabled = False
            txtAddress.Enabled = False
            End If
            
            If .OptOnAcct.Value = True Then
                OptOnAcct.Value = True
            Else
                OptCash.Value = True
            End If
            OptOnAcct.Value = True
            If .txtPayScheme.Text <> "" Then
                cboScheme.Text = .txtPayScheme.Text
                txtTerm.Text = .txtTerms.Text
                txtGives.Text = .txtNPays.Text
                txtTtlPrice.Text = .txtTotal.Text
                txtAmtDue.Text = .txtBillAmt.Text
                txtStartPay.Text = .txtPayStart.Text
                If Val(txtTtlPrice.Text) > 0 And Val(txtGives.Text) > 0 And Val(txtAmtDue.Text) = 0 Then Call txtGives_LostFocus
            Else
                'Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE (((DPLEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((DPLEDGER.STATUS)='OPEN'));"
                'mySQL
                Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE DPLEDGER.CUSTOMERID=" & txtCustomerID.Text & " AND DPLEDGER.STATUS='OPEN';"
                
                If CtrlNumExist(Qry) Then
                    chkDP.Value = 1
                    If txtPO.Text <> "" Then Label1(0).Caption = "O.R. NUMBER"
                End If
            End If
        End If
    End With
    OptOnAcct.Value = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsCustomer.State = adStateOpen Then rsCustomer.Close
End Sub

Private Sub OptCash_Click()
    If OptCash.Value = True Then
        FraTerms.Enabled = False
    Else
        FraTerms.Enabled = True
    End If
End Sub

Private Sub OptOnAcct_Click()
    If OptOnAcct.Value = True Then
        FraTerms.Enabled = True
    End If
End Sub

Private Sub OptOnAcct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboScheme.SetFocus
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtAmtDue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtStartPay.SetFocus
End Sub

Private Sub txtCusType_KeyPress(KeyAscii As Integer)
    If keysacii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtGives_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If keysacii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
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
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtPO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPO_LostFocus()
    Dim Qry As String
    
    If txtPO.Text = "" Then Exit Sub
    'Qry = "SELECT CUSTPO.*, CUSTPO.CPONO From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtPO.Text & "'));"
    'mySQL
    
    
    If UpDRef.Value > 0 Then
        Qry = "SELECT CUSTOMERNAME From CUSTPROMOCARDINFO WHERE CUSTOMERNO='" & txtPO.Text & "';"
        cboBranch.Text = GetFieldValue(Qry)
    Else
        Qry = "SELECT CUSTPO.*, CUSTPO.CPONO From CUSTPO WHERE CUSTPO.CPONO='" & txtPO.Text & "';"
        If CtrlNumExist(Qry) And Val(txtCustomerID.Text) = 67 Then Call ShowPO
    End If
    
End Sub

Private Sub ShowPO()
    Dim rsPO As New ADODB.Recordset
    Dim rsPOdtls As New ADODB.Recordset
    Dim Qry As String, myMsg As String, MyInv As String, myRem As String
    Dim POid As Integer
    
    myMsg = ""
    'Qry = "SELECT CUSTPO.* From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtPO.Text & "'));"
    'mySQL
    Qry = "SELECT CUSTPO.* From CUSTPO WHERE CUSTPO.CPONO='" & txtPO.Text & "';"
    
    With rsPO
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            POid = !CPOID
            cboBranch.Text = !CUSTOMER
            MyInv = IIf(IsNull(!InvoiceNO), "", !InvoiceNO)
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
                SPSorder.txtJOid.Text = IIf(IsNull(!JOBORDERID), "", !JOBORDERID)
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
        'Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, [PRICE]*[QTY] AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.STOREPRICE, SPAREPARTS.COST " & _
              "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE (((SPSORDERDETAIL.CPOID)=" & POid & "));"
        'mySQL
        Qry = "SELECT SPAREPARTS.QTYONHAND, SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.PRICE*SPSORDERDETAIL.QTY AS TTLPRICE, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.STOREPRICE, SPAREPARTS.COST " & _
              "FROM SPAREPARTS INNER JOIN SPSORDERDETAIL ON SPAREPARTS.SPAREPARTID = SPSORDERDETAIL.SPAREPARTID WHERE SPSORDERDETAIL.CPOID=" & POid & ";"
        
        rsPOdtls.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsPOdtls.RecordCount > 0 Then
            rsPOdtls.MoveFirst
            .flxOrd.Rows = 1
            Do While Not rsPOdtls.EOF
                .flxOrd.AddItem vbTab & rsPOdtls!QtyOnHand & vbTab & rsPOdtls!Qty & vbTab & IIf(IsNull(rsPOdtls!Unit), "", rsPOdtls!Unit) & vbTab & rsPOdtls!PartNumber & vbTab & rsPOdtls!Description & vbTab & vbTab & Format$(rsPOdtls!Price, "###,##0.00") & vbTab & Format$(rsPOdtls!TtlPrice, "###,##0.00") & vbTab & IIf(IsNull(rsPOdtls!Location), "", rsPOdtls!Location) & vbTab & rsPOdtls!SparepartID & vbTab & IIf(IsNull(rsPOdtls!ApplicModel), "", rsPOdtls!ApplicModel) & vbTab & vbTab & rsPOdtls!STOREPRICE & vbTab & rsPOdtls!Cost
                rsPOdtls.MoveNext
            Loop
        End If
        rsPOdtls.Close
        Call .ComputeTotal
        txtTtlPrice.Text = .txtTotal.Text
    End With
    'mySQL
    Qry = "SELECT ONACCTERMS.TERMS From ONACCTERMS WHERE ONACCTERMS.STATUS='Default' AND ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & ";"
    
    txtTerm.Text = GetFieldValue(Qry)
    cboScheme.Text = "MONTHLY"
    'txtGives.Text = 1
    Call txtGives_LostFocus
End Sub

Private Sub txtStartPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If keysacii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
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




Private Sub UpDRef_Change()
    If UpDRef.Value = 0 Then
        Label1(0).Caption = "P.O NUMBER"
        cboPointsType.Enabled = False
        txtPO.Text = ""
        cboPointsType.Text = ""
    Else
        Label1(0).Caption = "CARD NUMBER"
        cboPointsType.Enabled = True
    End If
End Sub


