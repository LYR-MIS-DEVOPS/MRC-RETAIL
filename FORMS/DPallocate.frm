VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form DPallocate 
   Caption         =   "Allocate Down Payment"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   5880
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   7560
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   18
      Tag             =   "XYF"
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5160
      Top             =   0
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   4320
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRefNo 
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboTrans 
      Height          =   315
      Left            =   600
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid flxDP 
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "     |<TRANSACTION         |^REF. NO.     |^  REF. DATE      |>    REF. AMOUNT   |> AMT. ALLOCATED |>O.R. BALANCE"
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   480
      Width           =   10335
      Begin VB.TextBox txtBalance 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   585
         Width           =   2175
      End
      Begin VB.TextBox txtORamt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   225
         Width           =   2175
      End
      Begin VB.TextBox txtORdate 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   945
         Width           =   2175
      End
      Begin VB.TextBox txtORno 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   585
         Width           =   2175
      End
      Begin VB.TextBox txtCustomer 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   225
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "BALANCE:"
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
         Index           =   4
         Left            =   6720
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "O.R. AMOUNT:"
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
         Index           =   3
         Left            =   6720
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "O.R. DATE:"
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
         Left            =   360
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "O.R. NUMBER:"
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
         Left            =   360
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "CUSTOMER:"
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
         Left            =   360
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "DPallocate.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Save"
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
      Image           =   "DPallocate.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "DPallocate.frx":0B34
      cBack           =   -2147483633
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL:"
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   6000
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   10800
      m_MDI_Height    =   6915
   End
   Begin VB.Menu mnuRow 
      Caption         =   "flxRow"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row            Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "DPallocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RowSrc As Integer, MySource As String
Dim ActRow As Integer
Dim ActCol As Integer

Private Sub cboTrans_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxDP.SetFocus
        flxDP.Col = 2
    End If
End Sub

Private Sub cboTrans_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxDP.TextMatrix(ActRow, 1) = cboTrans.Tag
        cboTrans.Tag = "skiplostfocus"
        cboTrans.Visible = False
        flxDP.SetFocus
        flxDP.Col = 1
    End If
End Sub

Private Sub cboTrans_LostFocus()
    If cboTrans.Tag = "skiplostfocus" Then
        cboTrans.Tag = ""
        Exit Sub
    End If
    flxDP.TextMatrix(ActRow, 1) = cboTrans.Text
    cboTrans.Visible = False
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim rsLedger As New ADODB.Recordset
    Dim Ans As Integer, i As Integer, Bal As Currency, MyJOid As Long
    Dim MyAlloc As String, MyInv As String, MyJob As String, vMsg As String, MyStat As String
    Dim eMsg As String, Qry As String, MyOR As String
    
    Select Case Index
        Case 0  'eXIT
            Unload Me
        Case 1  'Save
            vMsg = ValidateDP
            If vMsg <> "" Then
                MsgBox vMsg, vbExclamation + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save allocation of Down Payment?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            MyAlloc = ""
            MyInv = ""
            MyJob = ""
            With flxDP
                For i = 1 To .Rows - 1
                    MyAlloc = IIf(MyAlloc <> "", MyAlloc & ",", "") & .TextMatrix(i, 2) & "=" & Val(Replace(.TextMatrix(i, 5), ",", ""))
                    If .TextMatrix(i, 1) = "INVOICE" Then
                        MyInv = IIf(MyInv <> "", MyInv & ", ", "") & .TextMatrix(i, 2)
                    Else
                        MyJob = IIf(MyJob <> "", MyJob & ", ", "") & .TextMatrix(i, 2)
                    End If
                Next i
                Bal = Val(Replace(txtBalance.Text, ",", "")) - Val(Replace(txtTotal.Text, ",", ""))
                MyStat = IIf(Bal > 0, "OPEN", "CLOSED")
                'Qry = "SELECT DPLEDGER.AMTALLOC From DPLEDGER WHERE (((DPLEDGER.ORNO)='" & txtORno.Text & "'));"
                'mySQL
                Qry = "SELECT DPLEDGER.AMTALLOC From DPLEDGER WHERE DPLEDGER.ORNO='" & txtORNo.Text & "';"
                
                vMsg = GetFieldValue(Qry)
                MyAlloc = IIf(vMsg <> "", vMsg & ", ", "") & MyAlloc    'Add Previous to Existing allocation
                'Qry = "SELECT DPLEDGER.INVREF From DPLEDGER WHERE (((DPLEDGER.ORNO)='" & txtORno.Text & "'));"
                'mySQL
                Qry = "SELECT DPLEDGER.INVREF From DPLEDGER WHERE DPLEDGER.ORNO='" & txtORNo.Text & "';"
                
                vMsg = GetFieldValue(Qry)
                MyInv = IIf(vMsg <> "", vMsg & ", ", "") & MyInv        'Add Previous to Existing Invoices
                'Qry = "SELECT DPLEDGER.JOREF From DPLEDGER WHERE (((DPLEDGER.ORNO)='" & txtORno.Text & "'));"
                'mySQL
                Qry = "SELECT DPLEDGER.JOREF From DPLEDGER WHERE DPLEDGER.ORNO='" & txtORNo.Text & "';"
                
                vMsg = GetFieldValue(Qry)
                MyJob = IIf(vMsg <> "", vMsg & ", ", "") & MyJob        'Add Previous to Existing Job Orders
                'Update DPLedger
                eMsg = "Line 0065: Updating Down Payment Ledger"
                'LYRcon.Execute "UPDATE DPLEDGER SET DPLEDGER.AMTALLOC = '" & MyAlloc & "', DPLEDGER.STATUS = '" & MyStat & "', DPLEDGER.BALANCE = " & Bal & ", DPLEDGER.INVREF = '" & MyInv & "', DPLEDGER.JOREF = '" & MyJob & "' " & _
                               "WHERE (((DPLEDGER.ORNO)='" & txtORno.Text & "') AND ((DPLEDGER.CUSTOMERID)=" & Val(txtCustID.Text) & "));"
                'mySQL
                LYRcon.Execute "UPDATE DPLEDGER SET DPLEDGER.AMTALLOC = '" & MyAlloc & "', DPLEDGER.STATUS = '" & MyStat & "', DPLEDGER.BALANCE = " & Bal & ", DPLEDGER.INVREF = '" & MyInv & "', DPLEDGER.JOREF = '" & MyJob & "' " & _
                               "WHERE DPLEDGER.ORNO='" & txtORNo.Text & "' AND DPLEDGER.CUSTOMERID=" & Val(txtCustID.Text) & ";"
                
                'Update References
                For i = 1 To .Rows - 1
                    If InStr(.TextMatrix(i, 1), "INVOICE") > 0 Then
                        eMsg = "Line 0071: Updating Table Invoice field O.R. Number"
                        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.ORNO = '" & txtORno.Text & "' " & _
                                       "WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(i, 2) & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.ORNO = '" & txtORNo.Text & "' " & _
                                       "WHERE SPINVOICE.SPINVNO='" & .TextMatrix(i, 2) & "';"
                        
                        If Val(.TextMatrix(i, 8)) > 0 Then 'ReqID is available
                            'Qry = "SELECT REQUISITION.ORNO From Requisition WHERE (((REQUISITION.REQID)=" & Val(.TextMatrix(i, 8)) & "));"
                            'mySQL
                            Qry = "SELECT REQUISITION.ORNO From Requisition WHERE REQUISITION.REQID=" & Val(.TextMatrix(i, 8)) & ";"
                            
                            MyOR = GetFieldValue(Qry)
                            MyOR = IIf(MyOR <> "", MyOR & ", " & txtORNo.Text, txtORNo.Text)
                            eMsg = "Line 0078: Updating Table Requisition field O.R. Number"
                            'LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.ORNO = '" & MyOR & "' WHERE (((REQUISITION.REQID)=" & Val(.TextMatrix(i, 8)) & "));"
                            'mySQL
                            LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.ORNO = '" & MyOR & "' WHERE REQUISITION.REQID=" & Val(.TextMatrix(i, 8)) & ";"
                             
                        End If
                        If .TextMatrix(i, 9) = "CREDIT" Then 'Invoice Type is Credit
                            'old qry
                            'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS " & _
                                  "WHERE (((ONACCTERMS.SPINVNO)='" & .TextMatrix(i, 2) & "') AND ((ONACCTERMS.CUSTOMERID)=" & .TextMatrix(i, 7) & "));"
                            'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & .TextMatrix(i, 2) & "'));"
                            'mySQL
                            Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(i, 2) & "';"
                            
                            If CtrlNumExist(Qry) Then
                                Bal = Val(GetFieldValue(Qry))
                                Bal = Bal - Val(Replace(.TextMatrix(i, 5), ",", ""))
                                MyStat = IIf(Bal > 0, "OPEN", "CLOSED")
                                eMsg = "Line 0087: Updating Table On Account Terms"
                                'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal & ", ONACCTERMS.STATUS = '" & MyStat & "', ONACCTERMS.TTLAMTPAID = [ONACCTERMS].[TTLAMTPAID]+" & Val(Replace(.TextMatrix(i, 5), ",", "")) & _
                                        " WHERE (((ONACCTERMS.SPINVNO)='" & .TextMatrix(i, 2) & "') AND ((ONACCTERMS.CUSTOMERID)=" & .TextMatrix(i, 7) & "));"
                                'mySQL
                                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal & ", ONACCTERMS.STATUS = '" & MyStat & "', ONACCTERMS.TTLAMTPAID = ONACCTERMS.TTLAMTPAID+" & Val(Replace(.TextMatrix(i, 5), ",", "")) & _
                                        " WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(i, 2) & "' AND ONACCTERMS.CUSTOMERID=" & .TextMatrix(i, 7) & ";"
                                
                                'Update Ledger
                                eMsg = "Line 0092: Saving OR details to Ledger"
                                'rsLedger.Open "Ledger", LYRcon, adOpenKeyset, adLockOptimistic
                                
                                '**Access
                                'rsLedger.AddNew
                                'rsLedger!SPINVNO = .TextMatrix(i, 2)
                                'rsLedger!RefType = "Payment"
                                'rsLedger!RefNo = txtORno.Text
                                'rsLedger!RefDate = txtORdate.Text
                                'rsLedger!Credit = Val(Replace(.TextMatrix(i, 5), ",", ""))
                                ' old Qry = "SELECT LEDGER.OSBAL From Ledger Where (((Ledger.SPINVNO) = '" & .TextMatrix(i, 2) & "') And ((Ledger.CUSTOMERID) = " & .TextMatrix(i, 7) & ")) ORDER BY LEDGER.SEQNO DESC;"
                                'old Bal = Val(GetFieldValue(Qry))
                                'rsLedger!OSbal = Bal '- Val(Replace(.TextMatrix(i, 5), ",", ""))
                                'rsLedger!CUSTOMERID = Val(.TextMatrix(i, 7))
                                'rsLedger!SEQNO = MaxSeqNo(.TextMatrix(i, 2)) 'use function
                                'old If Val(.TextMatrix(i, 14)) > 0 Then rsLedger!JobID = Val(.TextMatrix(i, 14))
                                'rsLedger.Update
                                'rsLedger.Close
                                
                                'mySQL
                                LYRcon.Execute "Insert Into ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, CUSTOMERID, SEQNO)" & _
                                               "VALUES('" & .TextMatrix(i, 2) & "', 'Payment', '" & txtORNo.Text & "', '" & Format(txtORdate.Text, "yyyy-mm-dd") & "', " & _
                                               Val(Replace(.TextMatrix(i, 5), ",", "")) & ", " & Bal & ", " & Val(.TextMatrix(i, 7)) & ", " & MaxSeqNo(.TextMatrix(i, 2)) & ");"
                            End If
                        End If
                    Else    'Job Order
                        'Update Job Order OR No
                        'Qry = "SELECT SRVCJOBORDER.ORNO From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & .TextMatrix(i, 2) & "'));"
                        'mySQL
                        Qry = "SELECT SRVCJOBORDER.ORNO From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & .TextMatrix(i, 2) & "';"
                        
                        MyOR = GetFieldValue(Qry)
                        MyOR = IIf(MyOR <> "", MyOR & ", " & txtORNo.Text, txtORNo.Text)
                        'Qry = "SELECT SRVCJOBORDER.BALANCE From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & .TextMatrix(i, 2) & "'));"
                        'mySQL
                        Qry = "SELECT SRVCJOBORDER.BALANCE From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & .TextMatrix(i, 2) & "';"
                        
                        Bal = Val(GetFieldValue(Qry))
                        Bal = Bal - Val(Replace(.TextMatrix(i, 5), ",", ""))
                        eMsg = "Line 0117: Updating Table JobOrder field O.R. Number"
                        
                        'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Bal & ", SRVCJOBORDER.PAYMENTDONE = " & IIf(Bal <= 0, True, False) & ", SRVCJOBORDER.ORNO = '" & MyOR & "' " & _
                            "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & .TextMatrix(i, 2) & "'));"
                        
                        'mySQL
                        LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Bal & ", SRVCJOBORDER.PAYMENTDONE = " & IIf(Bal <= 0, 1, 0) & ", SRVCJOBORDER.ORNO = '" & MyOR & "' " & _
                            "WHERE SRVCJOBORDER.JOBORDERNO='" & .TextMatrix(i, 2) & "';"
                       
                        If .TextMatrix(i, 9) = "ON ACCOUNT" Then 'JOB ORDER Pay Type is On Account
                            'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS " & _
                                  " WHERE (((ONACCTERMS.JOBORDERNO)='" & .TextMatrix(i, 2) & "') AND ((ONACCTERMS.CUSTOMERID)=" & .TextMatrix(i, 7) & ") AND ((ONACCTERMS.SPINVNO) Is Null));"
                            'mySQL
                            Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS " & _
                                  " WHERE (((ONACCTERMS.JOBORDERNO)='" & .TextMatrix(i, 2) & "') AND ((ONACCTERMS.CUSTOMERID)=" & .TextMatrix(i, 7) & ") AND ((ONACCTERMS.SPINVNO) Is Null));"
                            
                            Bal = Val(GetFieldValue(Qry))
                            Bal = Bal - Val(Replace(.TextMatrix(i, 5), ",", ""))
                            MyStat = IIf(Bal > 0, "OPEN", "CLOSED")
                            eMsg = "Line 0126: Updating Table On Account Terms"
                            'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal & ", ONACCTERMS.STATUS = '" & MyStat & "', ONACCTERMS.TTLAMTPAID = [ONACCTERMS].[TTLAMTPAID]+" & Val(Replace(.TextMatrix(i, 5), ",", "")) & _
                                    " WHERE (((ONACCTERMS.JOBORDERNO)='" & .TextMatrix(i, 2) & "') AND ((ONACCTERMS.CUSTOMERID)=" & .TextMatrix(i, 7) & ") AND ((ONACCTERMS.SPINVNO) Is Null));"
                            'mySQl
                            LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal & ", ONACCTERMS.STATUS = '" & MyStat & "', ONACCTERMS.TTLAMTPAID = ONACCTERMS.TTLAMTPAID+" & Val(Replace(.TextMatrix(i, 5), ",", "")) & _
                                    " WHERE (((ONACCTERMS.JOBORDERNO)='" & .TextMatrix(i, 2) & "') AND ((ONACCTERMS.CUSTOMERID)=" & .TextMatrix(i, 7) & ") AND ((ONACCTERMS.SPINVNO) Is Null));"
                            'Update Ledger
                            eMsg = "Line 0092: Saving OR details to Ledger"
                            
                            'mySQL
                            MxSq = MaxSeqNo(.TextMatrix(i, 2))
                            Qry = "SELECT SRVCJOBORDER.JOBORDERID From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & .TextMatrix(i, 2) & "'));"
                            MyJOid = Val(GetFieldValue(Qry))
                            
                            LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, CUSTOMERID, SEQNO, JOBID) " & _
                                "VALUES('" & .TextMatrix(i, 2) & "', 'Payment', '" & txtORNo.Text & "', '" & Format(txtORdate.Text, "yyyy-mm-dd") & "', " & Val(Replace(.TextMatrix(i, 5), ",", "")) & ", " & Bal & ", " & Val(.TextMatrix(i, 7)) & ", " & MxSq & ", " & MyJOid & ")"
                            
                            'rsLedger.Open "Ledger", LYRcon, adOpenKeyset, adLockOptimistic
                            'rsLedger.AddNew
                            'rsLedger!SPINVNO = .TextMatrix(i, 2)
                            'rsLedger!RefType = "Payment"
                            'rsLedger!RefNo = txtORno.Text
                            'rsLedger!RefDate = txtORdate.Text
                            'rsLedger!Credit = Val(Replace(.TextMatrix(i, 5), ",", ""))
                            'old Qry = "SELECT LEDGER.OSBAL From Ledger Where (((Ledger.SPINVNO) = '" & .TextMatrix(i, 2) & "') And ((Ledger.CUSTOMERID) = " & .TextMatrix(i, 7) & ")) ORDER BY LEDGER.SEQNO DESC;"
                            'old Bal = Val(GetFieldValue(Qry))
                            'rsLedger!OSbal = Bal ' - Val(Replace(.TextMatrix(i, 5), ",", ""))
                            'rsLedger!CUSTOMERID = Val(.TextMatrix(i, 7))
                            'rsLedger!SEQNO = MaxSeqNo(.TextMatrix(i, 2)) 'use function
                            'Qry = "SELECT SRVCJOBORDER.JOBORDERID From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & .TextMatrix(i, 2) & "'));"
                            'rsLedger!JOBID = Val(GetFieldValue(Qry))
                            'rsLedger.Update
                            'rsLedger.Close
                        End If
                    End If
                Next i
            End With
            Bal = Val(Replace(txtBalance.Text, ",", "")) - Val(Replace(txtTotal.Text, ",", ""))
            MyStat = IIf(Bal > 0, "OPEN", "CLOSED")
            With DPSummary.flxDP
                .TextMatrix(RowSrc, 6) = Format(Bal, "#,##0.00")
                .TextMatrix(RowSrc, 5) = MyStat
                .TextMatrix(RowSrc, 7) = MyInv
                .TextMatrix(RowSrc, 8) = MyJob
                .TextMatrix(RowSrc, 11) = MyAlloc
                'Change row color if mystat=closed
                If MyStat = "CLOSED" Then
                    .Col = 1: .Row = RowSrc: .ColSel = 9: .RowSel = RowSrc: .FillStyle = flexFillRepeat: .CellBackColor = vbWhite
                    .Col = 2
                End If
            End With
            MsgBox "Down Payment successfully allocated!", vbInformation + vbOKOnly, "Save"
            Unload Me
        Case 2  'Cancel
            Ans = MsgBox("Do you want to Cancel Allocation of Down Payment?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            Unload Me
    End Select
End Sub

Function ValidateDP() As String
    Dim eMsg As String, Ttl As Currency
    
    eMsg = ""
    With flxDP
        If .TextMatrix(.Rows - 1, 1) = "" And .TextMatrix(.Rows - 1, 2) = "" And .TextMatrix(.Rows - 1, 5) = "" Then .RemoveItem .Rows - 1
        If .TextMatrix(.Rows - 1, 1) = "" Or .TextMatrix(.Rows - 1, 2) = "" Or .TextMatrix(.Rows - 1, 5) = "" Then
            eMsg = "Incomplete entry on last row of grid."
        End If
        Ttl = 0
        For i = 1 To .Rows - 1
            Ttl = Ttl + Val(Replace(.TextMatrix(i, 5), ",", ""))
        Next i
        If Ttl > Val(Replace(txtBalance.Text, ",", "")) Then
            eMsg = eMsg & vbCrLf & "Total Allocated Amount cannot be greater than O.R. Balance"
        End If
    End With
    ValidateDP = eMsg
End Function

Function MaxSeqNo(MyInv As String) As Integer
    Dim rsSeq As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE (((LEDGER.SPINVNO)='" & MyInv & "'));"
    'mySQL
    Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE LEDGER.SPINVNO='" & MyInv & "';"

    With rsSeq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MaxSeqNo = IIf(IsNull(.Fields(0)), 0, .Fields(0)) + 1
        Else
            MaxSeqNo = 1
        End If
        .Close
    End With
End Function

Private Sub flxDP_Click()
    ActRow = flxDP.Row
    ActCol = flxDP.Col
    
    If ActCol = 0 Or ActRow = 0 Or ActCol = 3 Or ActCol = 4 Or ActCol = 6 Or MySource = "SHOW" Then Exit Sub
    Select Case ActCol
    Case 1
        With cboTrans
            If flxDP.TextMatrix(ActRow, 1) <> "" Then .Tag = flxDP.TextMatrix(ActRow, 1)
            .Top = flxDP.CellTop + flxDP.Top
            .Left = flxDP.CellLeft + flxDP.Left
            .Width = flxDP.CellWidth
            .Text = flxDP.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 2
        With txtRefNo
            If flxDP.TextMatrix(ActRow, 2) <> "" Then .Tag = flxDP.TextMatrix(ActRow, 2) & "|" & flxDP.TextMatrix(ActRow, 3) & "|" & flxDP.TextMatrix(ActRow, 4) & "|" & flxDP.TextMatrix(ActRow, 5) & "|" & flxDP.TextMatrix(ActRow, 6) & "|" & flxDP.TextMatrix(ActRow, 7) & "|" & flxDP.TextMatrix(ActRow, 8) & "|" & flxDP.TextMatrix(ActRow, 9)
            .Top = flxDP.CellTop + flxDP.Top
            .Left = flxDP.CellLeft + flxDP.Left
            .Width = flxDP.CellWidth
            .Text = flxDP.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    Case 5
        With txtAmount
            If flxDP.TextMatrix(ActRow, 5) <> "" Then .Tag = flxDP.TextMatrix(ActRow, 5)
            .Top = flxDP.CellTop + flxDP.Top
            .Left = flxDP.CellLeft + flxDP.Left
            .Width = flxDP.CellWidth
            .Text = flxDP.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End Select

End Sub

Private Sub flxDP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxDP_Click
End Sub

Private Sub flxDP_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If ActCol = 0 Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuRow)
        End If
    End If
End Sub

Private Sub Form_Load()
    cboTrans.AddItem "CREDIT INVOICE"
    cboTrans.AddItem "CASH INVOICE"
    cboTrans.AddItem "JOB ORDER"
    Timer1.Enabled = True
End Sub

Private Sub mnuAdd_Click()
    If flxDP.TextMatrix(flxDP.Rows - 1, 1) = "" Or flxDP.TextMatrix(flxDP.Rows - 1, 2) = "" Or flxDP.TextMatrix(flxDP.Rows - 1, 5) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    Else
        flxDP.AddItem ""
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxDP.Rows = 2 And ActRow = 1 Then
            flxDP.Rows = 1
            flxDP.Rows = 2
        Else
            flxDP.RemoveItem ActRow
        End If
        Call ComputeTotals
    End If
End Sub

Private Sub ComputeTotals()
    Dim Ttl As Currency, i As Integer
    
    With flxDP
        Ttl = 0
        For i = 1 To .Rows - 1
            Ttl = Ttl + Val(Replace(.TextMatrix(i, 5), ",", ""))
        Next i
    End With
    txtTotal.Text = Format(Ttl, "#,##0.00")
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    With flxDP
        '1-TRANSACTION, 2-REF. NO., 3-REF. DATE, 4-REF. AMOUNT, 5-AMT. ALLOCATED, 6-O.R. BALANCE, 7-CUSTOMERID, 8-REQID, 9 - REF TYPE
        .ColWidth(0) = 300
        For i = 1 To 6
            .ColWidth(i) = (.Width - 500) / 6
        Next i
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        txtTotal.Width = .ColWidth(5)
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4)
        lblTotal.Left = txtTotal.Left - lblTotal.Width
    End With
End Sub

Private Sub Timer1_Timer()
    Dim rsInv As New ADODB.Recordset
    Dim MyJobs() As String, MyInvoices() As String, MyAA() As String, MyBB As String
    Dim i As Integer, InvDate As Date, InvAmt As Currency, AmtAlloc As Currency, Ttl As Currency
    Dim Qry As String, j As Integer
    
    With DPSummary.flxDP
        txtCustomer.Text = .TextMatrix(RowSrc, 3)
        txtORNo.Text = .TextMatrix(RowSrc, 2)
        txtORdate.Text = .TextMatrix(RowSrc, 1)
        txtORamt.Text = .TextMatrix(RowSrc, 4)
        txtBalance.Text = .TextMatrix(RowSrc, 6)
        txtCustID.Text = .TextMatrix(RowSrc, 10)
        If MySource = "SHOW" Then
            Me.Caption = "Allocation of Down Payment"
            cmdTrans(0).Visible = True
            cmdTrans(1).Visible = False
            cmdTrans(2).Visible = False
            If .TextMatrix(RowSrc, 7) <> "" Then    'Invoice available
                MyInvoices = Split(.TextMatrix(RowSrc, 7), ",")
                flxDP.Rows = 1
                For i = LBound(MyInvoices) To UBound(MyInvoices)
                    'Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & Trim(MyInvoices(i)) & "'));"
                    'mySQL
                    Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & Trim(MyInvoices(i)) & "';"
                    rsInv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If rsInv.RecordCount > 0 Then
                        rsInv.MoveFirst
                        InvDate = rsInv!InvDate
                        InvAmt = rsInv!Amount
                        If InStr(.TextMatrix(RowSrc, 11), Trim(MyInvoices(i))) > 0 Then
                            If InStr(.TextMatrix(RowSrc, 11), ",") > 0 Then
                                MyAA = Split(.TextMatrix(RowSrc, 11), ",")
                                For j = LBound(MyAA) To UBound(MyAA)
                                    If InStr(MyAA(j), Trim(MyInvoices(i))) > 0 Then
                                        MyBB = MyAA(j)
                                        Exit For
                                    End If
                                Next j
                            Else
                                MyBB = .TextMatrix(RowSrc, 11)
                            End If
                            AmtAlloc = Val(Replace(MyBB, Trim(MyInvoices(i)) & "=", ""))
                        Else
                            If Val(.TextMatrix(RowSrc, 4)) > InvAmt Then
                                AmtAlloc = InvAmt
                            Else
                                AmtAlloc = Val(Replace(.TextMatrix(RowSrc, 4), ",", ""))
                            End If
                        End If
                        flxDP.AddItem vbTab & "INVOICE" & vbTab & Trim(MyInvoices(i)) & vbTab & Format(InvDate, "mm/dd/yyyy") & vbTab & Format(InvAmt, "#,##0.00") & vbTab & Format(AmtAlloc, "#,##0.00") & vbTab & Format(Val(Replace(txtORamt.Text, ",", "")) - AmtAlloc, "#,##0.00")
                    End If
                    rsInv.Close
                Next i
            Else
                flxDP.Rows = 1
            End If
            If .TextMatrix(RowSrc, 8) <> "" Then 'Job Order available
                MyJobs = Split(.TextMatrix(RowSrc, 8), ",")
                'flxDP.Rows = 1
                For i = LBound(MyJobs) To UBound(MyJobs)
                    'Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & Trim(MyJobs(i)) & "'));"
                    'mySQL
                    Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & Trim(MyJobs(i)) & "';"
                    rsInv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If rsInv.RecordCount > 0 Then
                        rsInv.MoveFirst
                        InvDate = rsInv!DateReceived
                        If rsInv!DISCOUNT > 0 Then
                            InvAmt = rsInv!TotalCost - (rsInv!TotalCost * (rsInv!DISCOUNT / 100))
                        Else
                            InvAmt = rsInv!TotalCost
                        End If
                        If InStr(.TextMatrix(RowSrc, 11), Trim(MyJobs(i))) > 0 Then
                            If InStr(.TextMatrix(RowSrc, 11), ",") > 0 Then
                                MyAA = Split(.TextMatrix(RowSrc, 11), ",")
                                For j = LBound(MyAA) To UBound(MyAA)
                                    If InStr(MyAA(j), Trim(MyJobs(i))) > 0 Then
                                        MyBB = MyAA(j)
                                        Exit For
                                    End If
                                Next j
                            Else
                                MyBB = .TextMatrix(RowSrc, 11)
                            End If
                            AmtAlloc = Val(Replace(MyBB, Trim(MyJobs(i)) & "=", ""))
                        Else
                            If Val(.TextMatrix(RowSrc, 4)) > InvAmt Then    'OR amount> Job Cost
                                AmtAlloc = InvAmt
                            Else
                                AmtAlloc = Val(Replace(.TextMatrix(RowSrc, 4), ",", ""))
                            End If
                        End If
                        flxDP.AddItem vbTab & "JOB ORDER" & vbTab & MyJobs(i) & vbTab & Format(InvDate, "mm/dd/yyyy") & vbTab & Format(InvAmt, "#,##0.00") & vbTab & Format(AmtAlloc, "#,##0.00") & vbTab & Format(Val(Replace(txtORamt.Text, ",", "")) - AmtAlloc, "#,##0.00")
                    End If
                    rsInv.Close
                Next i
            End If
        ElseIf MySource = "ALLOCATE" Then
            cmdTrans(0).Visible = False
            cmdTrans(1).Visible = True
            cmdTrans(2).Visible = True
            cmdTrans(1).Left = 0
            cmdTrans(2).Left = cmdTrans(1).Width
        End If
    End With
    If MySource = "SHOW" Then
        Ttl = 0
        For i = 1 To flxDP.Rows - 1
            Ttl = Ttl + Val(Replace(flxDP.TextMatrix(i, 5), ",", ""))
            flxDP.TextMatrix(i, 6) = Val(Replace(txtORamt.Text, ",", "")) - Ttl
        Next i
        txtTotal.Text = Format(Ttl, "#,##0.00")
    End If
    Timer1.Enabled = False
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxDP.SetFocus
        flxDP.Row = flxDP.Rows - 1
        flxDP.Col = 0
    End If
End Sub

Private Sub txtAmount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxDP.TextMatrix(ActRow, 5) = txtAmount.Tag
        txtAmount.Tag = "skiplostfocus"
        txtAmount.Visible = False
        flxDP.SetFocus
        flxDP.Col = 5
    End If
End Sub

Private Sub txtAmount_LostFocus()
    Dim Ttl As Currency, i As Integer, eMsg As String
    
    If txtAmount.Tag = "skiplostfocus" Then
        txtAmount.Tag = ""
        Exit Sub
    End If
    'Validate amount if > Ref. Amount
    If Val(Replace(txtAmount.Text, ",", "")) > Val(Replace(flxDP.TextMatrix(ActRow, 4), ",", "")) Then
        eMsg = IIf(flxDP.TextMatrix(ActRow, 4) <> "", "Allocated amount must not be greater than " & flxDP.TextMatrix(ActRow, 1) & " amount", "No available " & flxDP.TextMatrix(ActRow, 1) & " amount")
        MsgBox eMsg, vbExclamation + vbOKOnly, "Validation Error"
        txtAmount.Visible = False
        Exit Sub
    End If
    'Get Total
    Ttl = 0
    For i = 1 To flxDP.Rows - 1
        If i <> ActRow Then
            Ttl = Ttl + Val(Replace(flxDP.TextMatrix(i, 5), ",", ""))
        End If
    Next i
    If Ttl + Val(txtAmount.Text) > Val(Replace(txtBalance.Text, ",", "")) Then
        MsgBox "Total allocated amount cannot be greater than O.R. Balance"
        txtAmount.Visible = False
        Exit Sub
    End If
    flxDP.TextMatrix(ActRow, 5) = Format(txtAmount.Text, "#,##0.00")
    flxDP.TextMatrix(ActRow, 6) = Format(Val(Replace(txtBalance.Text, ",", "")) - (Ttl + Val(txtAmount.Text)), "#,##0.00")
    txtAmount.Visible = False
    Call ComputeTotals
    If flxDP.TextMatrix(flxDP.Rows - 1, 1) <> "" And flxDP.TextMatrix(flxDP.Rows - 1, 2) <> "" And flxDP.TextMatrix(flxDP.Rows - 1, 3) <> "" And flxDP.TextMatrix(flxDP.Rows - 1, 5) <> "" Then flxDP.AddItem ""
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxDP.SetFocus
        flxDP.Col = 5
    End If
End Sub

Private Sub txtRefNo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim PrevInv() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        If txtRefNo.Tag <> "" Then
            PrevInv = Split(txtRefNo.Tag, "|")
            For i = LBound(PrevInv) To UBound(PrevInv)
                flxDP.TextMatrix(ActRow, i + 2) = PrevInv(i)
            Next i
        End If
        txtRefNo.Tag = "skiplostfocus"
        txtRefNo.Visible = False
        flxDP.SetFocus
        flxDP.Col = 3
    End If
End Sub

Private Sub txtrefno_LostFocus()
    Dim rsInv As New ADODB.Recordset
    Dim rsJob As New ADODB.Recordset
    Dim Qry As String, MyCust As String, Ans As Integer
    Dim InvAmt As Currency, RetAmt As Currency
    
    If txtRefNo.Tag = "skiplostfocus" Then
        txtRefNo.Tag = ""
        Exit Sub
    End If
    If txtRefNo.Text = "" Then txtRefNo.Visible = False: Exit Sub
    If flxDP.TextMatrix(ActRow, 1) = "" Then
        MsgBox "Please select type of transaction in first column!", vbExclamation + vbOKOnly, "Unknown Transaction"
        txtRefNo.Visible = False
        Exit Sub
    Else
        If InStr(flxDP.TextMatrix(ActRow, 1), "INVOICE") > 0 Then
            'Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
            'mySQL
            If flxDP.TextMatrix(ActRow, 1) = "CASH INVOICE" Then
                Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "' AND SPINVOICE.INVTYPE='CASH';"
            ElseIf flxDP.TextMatrix(ActRow, 1) = "CREDIT INVOICE" Then
                Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "' AND SPINVOICE.INVTYPE='CREDIT';"
            Else
                Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "';"
            End If
            With rsInv
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    flxDP.TextMatrix(ActRow, 3) = Format(!InvDate, "mm/dd/yyyy")
                    If !InvType = "CASH" Then
                        InvAmt = !Amount
                        'Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                               "WHERE (((SRETURNSP.INVOICE)='" & txtRefNo.Text & "') AND ((SRETURNSP.REMARKS) Not Like 'CANCELLED'));"
                        'mySQL
                        Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                               "WHERE (((SRETURNSP.INVOICE)='" & txtRefNo.Text & "') AND ((SRETURNSP.REMARKS) Not Like 'CANCELLED'));"
                        
                        RetAmt = Val(GetFieldValue(Qry))
                        flxDP.TextMatrix(ActRow, 4) = Format(InvAmt - RetAmt, "#,##0.00")
                    Else
                        'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & txtRefNo.Text & "'));"
                        'mySQL
                        Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & txtRefNo.Text & "';"
                        
                        If CtrlNumExist(Qry) Then
                            InvAmt = Val(GetFieldValue(Qry))
                            flxDP.TextMatrix(ActRow, 4) = Format(InvAmt, "#,##0.00")
                        Else
                            InvAmt = !Amount
                            'Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                                   "WHERE (((SRETURNSP.INVOICE)='" & txtRefNo.Text & "') AND ((SRETURNSP.REMARKS) Not Like 'CANCELLED'));"
                            'mySQL
                            Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                                   "WHERE (((SRETURNSP.INVOICE)='" & txtRefNo.Text & "') AND ((SRETURNSP.REMARKS) Not Like 'CANCELLED'));"
                            
                            RetAmt = Val(GetFieldValue(Qry))
                            flxDP.TextMatrix(ActRow, 4) = Format(InvAmt - RetAmt, "#,##0.00")
                        End If
                    End If
                    MyCust = !chargeTo
                    flxDP.TextMatrix(ActRow, 7) = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
                    flxDP.TextMatrix(ActRow, 8) = IIf(IsNull(!ReqID), 0, !ReqID)
                    flxDP.TextMatrix(ActRow, 9) = !InvType
                End If
                .Close
            End With
            If MyCust <> txtCustomer.Text Then
                Ans = MsgBox("Invoice was charge to " & MyCust & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, "Different Customer")
                If Ans = vbNo Then
                    flxDP.TextMatrix(ActRow, 3) = ""
                    flxDP.TextMatrix(ActRow, 4) = ""
                    flxDP.TextMatrix(ActRow, 7) = ""
                    flxDP.TextMatrix(ActRow, 8) = ""
                    flxDP.TextMatrix(ActRow, 9) = ""
                    flxDP.TextMatrix(ActRow, 2) = ""
                    txtRefNo.Visible = False
                    Exit Sub
                End If
            End If
        ElseIf flxDP.TextMatrix(ActRow, 1) = "JOB ORDER" Then
            'Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
            With rsJob
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    flxDP.TextMatrix(ActRow, 3) = Format(!DatePromised, "mm/dd/yyyy")
                    flxDP.TextMatrix(ActRow, 4) = Format(!Balance, "#,##0.00")
                    MyCust = !CUSTOMER
                    flxDP.TextMatrix(ActRow, 7) = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
                    flxDP.TextMatrix(ActRow, 8) = ""
                    flxDP.TextMatrix(ActRow, 9) = !PayType
                End If
                .Close
            End With
            If MyCust <> txtCustomer.Text Then
                Ans = MsgBox("Job Order was charge to " & MyCust & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, "Different Customer")
                If Ans = vbNo Then
                    flxDP.TextMatrix(ActRow, 3) = ""
                    flxDP.TextMatrix(ActRow, 4) = ""
                    flxDP.TextMatrix(ActRow, 7) = ""
                    flxDP.TextMatrix(ActRow, 8) = ""
                    flxDP.TextMatrix(ActRow, 9) = ""
                    flxDP.TextMatrix(ActRow, 2) = ""
                    txtRefNo.Visible = False
                    Exit Sub
                End If
            End If
        End If
    End If
    If flxDP.TextMatrix(ActRow, 3) = "" Then
        MsgBox flxDP.TextMatrix(ActRow, 1) & " NO. " & txtRefNo.Text & " Not Found!", vbExclamation + vbOKOnly, "Validation Error"
    Else
        flxDP.TextMatrix(ActRow, 2) = txtRefNo.Text
        txtRefNo.Visible = False
    End If
End Sub
