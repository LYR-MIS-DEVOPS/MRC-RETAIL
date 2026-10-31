VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form InvDetails 
   Caption         =   "Invoice Details"
   ClientHeight    =   4365
   ClientLeft      =   5550
   ClientTop       =   4335
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7905
   Begin VB.TextBox txtPrintCopy 
      Height          =   285
      Left            =   3840
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrCancel 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4560
      Top             =   4440
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Ok"
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
      Image           =   "InvDetails.frx":0000
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxInv 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   8
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      FormatString    =   "     |^QTY.|^UNIT|<PART NUMBER|<DESCRIPTION|>UNIT PRICE|>TOTAL PRICE "
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Tag             =   "XYF"
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Print"
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
      Image           =   "InvDetails.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.Label lblRemarks 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL:"
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
      Left            =   5040
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblInvDate 
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
      Left            =   6600
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblInvNo 
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
      Left            =   6600
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblAddress 
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
      Left            =   1560
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblChargeTo 
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
      Left            =   1560
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "INVOICE DATE:"
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
      Left            =   5160
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "INVOICE NO."
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
      Left            =   5160
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ADDRESS:"
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
      TabIndex        =   2
      Tag             =   "XYF"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "CHARGE TO:"
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
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   120
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   240
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   8145
      m_MDI_Height    =   4935
   End
End
Attribute VB_Name = "InvDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyInv As String, CancelInv As Boolean, iRow As Integer, MyType As String

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim Ans As Integer, i As Integer, QHnd As Integer
    Dim Qry As String
    
    Ans = MsgBox("Do you want to Print Invoice?", vbQuestion + vbYesNo, "Print")
    If Ans = vbNo Then Exit Sub
    If Val(txtPrintCopy.Text) = 0 Then
        'Update SPITR
        'LYRcon.Execute "UPDATE SPITR SET SPITR.REMARKS = '" & Left$(Me.Caption, Len(Me.Caption) - 8) & "' " & _
                       "WHERE (((SPITR.REFERENCE)='SI#-" & lblInvNo.Caption & "') AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)='Temp'));"
        
        'mySQL
        LYRcon.Execute "UPDATE SPITR SET SPITR.REMARKS = '" & Left(Me.Caption, Len(Me.Caption) - 8) & "' " & _
                       "WHERE SPITR.REFERENCE='SI#-" & lblInvNo.Caption & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='Temp';"
        
        'Update Print Copy
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.PRINTCOPY = [SPINVOICE].[PRINTCOPY]+ 1 WHERE (((SPINVOICE.SPINVNO)='" & lblInvNo.Caption & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.PRINTCOPY = SPINVOICE.PRINTCOPY+ 1 WHERE SPINVOICE.SPINVNO='" & lblInvNo.Caption & "';"
        
        'Update Qty On hand
        For i = 1 To flxInv.Rows - 1
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND =  [SPAREPARTS].[QTYONHAND] -" & Val(flxInv.TextMatrix(i, 1)) & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = #" & Now & "# " & _
                           "WHERE (((SPAREPARTS.SPAREPARTID)=" & flxInv.TextMatrix(i, 7) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND =  SPAREPARTS.QTYONHAND -" & Val(flxInv.TextMatrix(i, 1)) & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                           "WHERE SPAREPARTS.SPAREPARTID=" & flxInv.TextMatrix(i, 7) & ";"
            
        Next i
    End If
    'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='SI#-" & MyInv & "'));"
    
    'mySQL
'    Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE " & _
'          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='SI#-" & MyInv & "';"
'
    Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.`OUT` AS TTLPRICE, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='SI#-" & MyInv & "' AND SPITR.REMARKS = '" & MyType & "' ORDER BY SPAREPARTS.PARTNUMBER;"
    
    PRNmc.MySource = "SPINVOICE"
    PRNmc.QryPrn = Qry
    PRNmc.ParamValue1 = Me.Caption
    PRNmc.ParamValue2 = lblChargeTo.Caption
    PRNmc.ParamValue3 = lblAddress.Caption
    PRNmc.ParamValue4 = lblInvNo.Caption
    PRNmc.ParamValue5 = lblInvDate.Caption
    PRNmc.ParamValue6 = "RE-PRINT COPY " & lblRemarks.Caption
    Unload Me
    PRNmc.Show
End Sub

Private Sub Form_Load()
    Dim rsInv As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, MyUnit As String, MyQty As Integer, MyOR As String, MyRef As String
    
    'Qry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE, SPINVOICE.CUSTOMERID, SPINVOICE.ORNO, SPINVOICE.PRINTCOPY " & _
          "From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & MyInv & "'));"
    'mySQL
    Qry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE, SPINVOICE.CUSTOMERID, SPINVOICE.ORNO, SPINVOICE.PRINTCOPY " & _
          "From SPINVOICE WHERE SPINVOICE.SPINVNO='" & MyInv & "' AND SPINVOICE.INVTYPE = '" & MyType & "';"
    
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            lblChargeTo.Caption = !ChargeTo
            lblAddress.Caption = !Address
            lblInvNo.Caption = !SPINVNO
            lblInvDate.Caption = !InvDate
            lblTotal.Caption = Format(!Amount, "#,##0.00")
            Me.Caption = !InvType & " INVOICE"
            txtCustID.Text = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
            lblRemarks.Caption = IIf(IsNull(!Remarks), "", !Remarks)
            MyOR = IIf(IsNull(!ORNO), "", !ORNO)
            txtPrintCopy.Text = IIf(IsNull(!PRINTCOPY), 0, !PRINTCOPY)
        End If
        .Close
    End With
    
    If lblInvNo.Caption = "" Then Exit Sub
    'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [PRICE]*[OUT] AS TTLPRICE, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='SI#-" & MyInv & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    If InStr(MyInv, "TRA#:") = 0 Then
        MyRef = "SI#-" & MyInv
    Else
        MyRef = MyInv
    End If
    Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.`OUT` AS TTLPRICE, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "' AND SPITR.REMARKS = '" & MyType & "' ORDER BY SPAREPARTS.PARTNUMBER;"
    
    flxInv.Rows = 1
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If InStr(!Description, "ENGINE OIL") > 0 Or InStr(!Description, "OIL 4T") > 0 Then
                    MyUnit = IIf(!OUT > 1, "Qrts.", "Qrt.")
                Else
                    MyUnit = IIf(!OUT > 1, "Pcs.", "Pc.")
                End If
                flxInv.AddItem vbTab & !OUT & vbTab & MyUnit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!TtlPrice, "#,##0.00") & vbTab & !SparepartID
                MyQty = ReturnedQty(!SparepartID)
                If MyQty > 0 Then
                    flxInv.Col = 1: flxInv.Row = flxInv.Rows - 1: flxInv.ColSel = 7: flxInv.RowSel = flxInv.Rows - 1: flxInv.FillStyle = flexFillRepeat: flxInv.CellBackColor = RGB(250, 250, 180)
                    flxInv.Col = 2
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    If CancelInv And MyType = "CREDIT" Then
        cmdOK.Visible = False
        cmdPrint.Visible = False
        If Me.Caption = "CREDIT INVOICE" And MyOR <> "" Then
                MsgBox "Can not Cancel Invoice because an O.R. was already issued!", vbCritical + vbOKOnly, "CANCEL DENIED"
                cmdOK.Visible = True
                Exit Sub
        End If
        'Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.SPINVNO)='" & lblInvNo.Caption & "') AND ((LEDGER.REFTYPE)<>'CANCELLED'));"
        'mySQL
        
        Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.SPINVNO='" & lblInvNo.Caption & "' AND LEDGER.REFTYPE<>'CANCELLED' AND LEDGER.JOBID=0;"
        
        If CtrlNumExist(Qry) And Me.Caption = "CREDIT INVOICE" Then
                MsgBox "Can not Cancel Invoice because an O.R. was already issued!", vbCritical + vbOKOnly, "CANCEL DENIED"
                cmdOK.Visible = True
                Exit Sub
        End If
        tmrCancel.Enabled = True
    ElseIf CancelInv And MyType = "CASH" Then
        cmdOK.Visible = True
        tmrCancel.Enabled = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MyInv = ""
    CancelInv = False
End Sub

Function ReturnedQty(SPid As Long) As Integer
    Dim rsRetQ As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.SPAREPARTID)=" & SPid & ") AND ((SPITR.PACKINGLIST)='" & MyInv & "'));"
    'mySQL
    Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.SPAREPARTID=" & SPid & " AND SPITR.PACKINGLIST='" & MyInv & "';"
    
    With rsRetQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ReturnedQty = !RECEIVE
        Else
            ReturnedQty = 0
        End If
        .Close
    End With
End Function

Private Sub RXResizer1_ResizeComplete()
    With flxInv
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.1
        .ColWidth(2) = (.Width - 300) * 0.09
        .ColWidth(3) = (.Width - 300) * 0.25
        .ColWidth(4) = (.Width - 300) * 0.25
        .ColWidth(5) = (.Width - 300) * 0.15
        .ColWidth(6) = (.Width - 300) * 0.15
        .ColWidth(7) = 0
    End With

End Sub

Private Sub tmrCancel_Timer()
    Dim Ans As Integer, MyReason As String, Qry As String
    Dim myAcctDesc As String * 50
    Dim i As Integer, ReqID As Integer, POid As Integer

    If DateValue(lblInvDate.Caption) <> Date Then
        MsgBox "Cancel not allowed for this Invoice!", vbExclamation + vbOKOnly, "Validation Error"
        Unload Me
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Cancel " & Me.Caption & "?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then
        Unload Me
        Exit Sub
    End If
    MyReason = InputBox("Please indicate reason for Cancellation of " & Me.Caption & ":", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of " & Me.Caption & " failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update Tables
    With flxInv
        'Update table SPInvoice
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'CANCELLED', SPINVOICE.AMOUNT = 0 " & _
                       "WHERE (((SPINVOICE.SPINVNO)='" & MyInv & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'CANCELLED', SPINVOICE.AMOUNT = 0 " & _
                       "WHERE SPINVOICE.SPINVNO='" & MyInv & "' AND SPINVOICE.INVTYPE = '" & MyType & "';"
        
        If InStr(Me.Caption, "CREDIT") > 0 Then 'Invoice type = "Credit"
            'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'CANCELLED' WHERE (((ONACCTERMS.SPINVNO)='" & MyInv & "'));"
            'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = [CUSTOMER].[OSBAL]-" & PureNum(lblTotal.Caption) & _
            " WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
            
            'mySQL
            LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'CLOSED' WHERE ONACCTERMS.SPINVNO='" & MyInv & "' and ONACCTERMS.CUSTOMERID=" & txtCustID.Text & ";"
            LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL-" & Val(Replace(lblTotal.Caption, ",", "")) & _
            " WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
            
        Else
            'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.DOWNPAYMENT = 0, PAYMENTS.SPAREPARTS = 0, PAYMENTS.SERVICE = 0, PAYMENTS.REMARKS = 'CANCELLED' " & _
                           "WHERE (((PAYMENTS.INVOICENO)='" & MyInv & "'));"
        
            'mySQL
            LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.DOWNPAYMENT = 0, PAYMENTS.SPAREPARTS = 0, PAYMENTS.SERVICE = 0, PAYMENTS.REMARKS = 'CANCELLED' " & _
                           "WHERE PAYMENTS.INVOICENO='" & MyInv & "';"
        
        End If
        'Qry = "SELECT REQUISITION.REQID From Requisition WHERE (((REQUISITION.INVOICENO)='" & MyInv & "'));"
        'mySQL
        Qry = "SELECT REQUISITION.REQID From Requisition WHERE REQUISITION.INVOICENO='" & MyInv & "';"
        
        ReqID = Val(GetFieldValue(Qry))
        'If ReqID > 0 Then LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = Null WHERE (((REQUISITION.REQID)=" & ReqID & "));"
        'Qry = "SELECT CUSTPO.CPOID From CUSTPO WHERE (((CUSTPO.INVOICENO)='" & MyInv & "'));"
        'mySQL
        If ReqID > 0 Then LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = Null WHERE REQUISITION.REQID=" & ReqID & ";"
        Qry = "SELECT CUSTPO.CPOID From CUSTPO WHERE CUSTPO.INVOICENO='" & MyInv & "';"
        
        POid = Val(GetFieldValue(Qry))
        If POid > 0 Then LYRcon.Execute "UPDATE CUSTPO SET CUSTPO.INVOICENO = Null WHERE (((CUSTPO.CPOID)=" & POid & "));"

        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+" & Val(.TextMatrix(i, 1)) & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 7) & "));"
                'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.OUT = 0 " & _
                               "WHERE (((SPITR.REFERENCE)='SI#-" & MyInv & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 7) & "));"
                'mySQL
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+" & Val(.TextMatrix(i, 1)) & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 7) & ";"
                LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.`OUT` = 0 " & _
                               "WHERE SPITR.REFERENCE='SI#-" & MyInv & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 7) & " AND SPITR.REMARKS = '" & MyType & "';"
            
            Next i
        End If
        'LYRcon.Execute "INSERT INTO AUDITRANS (TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'SALES INVOICE', '" & MyInv & "', #" & lblInvDate.Caption & "#, '" & MyReason & "', " & UserID & ");"
        'mySQl
        LYRcon.Execute "INSERT INTO AUDITRANS (TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'SALES INVOICE', '" & MyInv & " " & MyType & "', '" & Format(lblInvDate.Caption, "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
    
    End With
    'Update Grid
    With SPSalesRpt.flxAcct
        .TextMatrix(iRow, 5) = 0
        .TextMatrix(iRow, 6) = "CANCELLED"
    End With
    MsgBox Me.Caption & " Cancelled!", vbInformation + vbOKOnly
    Unload Me
    tmrCancel.Enabled = False
End Sub
