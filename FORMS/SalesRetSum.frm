VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SalesRetSum 
   Caption         =   "Spare Parts Sales Return Summary"
   ClientHeight    =   8475
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxORDER 
      Height          =   7095
      Left            =   120
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   1350
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12515
      _Version        =   393216
      RowHeightMin    =   285
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   11655
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9600
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8280
         TabIndex        =   6
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
         Image           =   "SalesRetSum.frx":0000
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         ToolTipText     =   "Execute"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "SalesRetSum.frx":059A
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         ToolTipText     =   "Execute"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "SalesRetSum.frx":0B34
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   16
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   17
         ToolTipText     =   "Execute"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "SalesRetSum.frx":10CE
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
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
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Sales Return"
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
      Image           =   "SalesRetSum.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Refresh "
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
      Image           =   "SalesRetSum.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   10
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Summary"
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
      Image           =   "SalesRetSum.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   11
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
      Image           =   "SalesRetSum.frx":2736
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxCancel 
      Height          =   1695
      Left            =   1560
      TabIndex        =   18
      Top             =   8640
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "Qty|<Unit|<Part Number|<Description |<Price | Amount|<SPID"
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11985
      m_MDI_Height    =   8985
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Parts Sales Return Summary"
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
      Left            =   7200
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
      Caption         =   "Spare Parts Sales Return Summary"
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
      Left            =   7215
      TabIndex        =   3
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuSalRet 
      Caption         =   "Sales Return"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Sales Return"
      End
   End
End
Attribute VB_Name = "SalesRetSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSPord As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer

Private Sub RefreshRecord(dCode As Integer)
    Dim Qry As String, i As Integer
    
    If dCode = 1 Then
        Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS, SRETURNSP.PRINTCOPY, SRETURNSP.RETURNID " & _
              "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where (((SRETURNSP.RETDATE) = #" & dFrom.Value & "#)) ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO DESC;"
    Else
        Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS, SRETURNSP.PRINTCOPY, SRETURNSP.RETURNID " & _
              "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where (((SRETURNSP.RETDATE) >= #" & dFrom.Value & "# And (SRETURNSP.RETDATE) <= #" & dTo.Value & "#)) ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO DESC;"
    End If
    If rsSPord.State = adStateOpen Then rsSPord.Close
    rsSPord.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPord.RecordCount = 0 Then
        flxORDER.Clear
        flxORDER.Rows = 2
        flxORDER.Cols = 9
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxORDER.DataSource = rsSPord
    ActRow = 1
    flxORDER.Visible = True
    Call RXResizer1_ResizeComplete
End Sub


Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxORDER
        For i = fRow To .Rows - 1
            For j = fCol To 8
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 8:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 8, i + 1, i)
                    fCol = IIf(j = 8, 1, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
End Sub

Private Sub cmdNext_Click()
    cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    dTo.Visible = True
    cmdRefresh.Visible = True
    dTo.Value = dFrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    dTo.Visible = False
    cmdRefresh.Visible = False
    Call RefreshRecord(1)
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshRecord(2)
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    
    Select Case Index
        Case 0  'Show Packing No
            If flxORDER.Rows = 2 Then ActRow = 1
            If rsSPord.RecordCount > 0 And ActRow <> 0 Then
                'SalesReturn.Show
                Call SalesReturn.DisplaySalesReturn(flxORDER.TextMatrix(ActRow, 9))
                Unload Me
            Else
                SalesReturn.Show
                Unload Me
            End If
        Case 1  'Refresh
            Call RefreshRecord(Val(cmdNext.Tag))
        Case 2  'Print
        
        Case 3  'Exit
            Unload Me
            MotoRep.tbMain.Visible = True
    End Select
        
End Sub

Private Sub dFrom_CloseUp()
    If cmdNext.Tag = "1" Then
        Call RefreshRecord(1)
    End If
End Sub

Private Sub dTo_CloseUp()
    If dTo.Value < dFrom.Value Then
        MsgBox "Please choose a date not less than " & dFrom.Value, vbExclamation + vbOKOnly, "Validation Error"
        dTo.Value = dFrom.Value
    End If
End Sub


Private Sub flxOrder_Click()
    ActRow = flxORDER.Row
End Sub

Private Sub flxOrder_EnterCell()
    ActRow = flxORDER.Row
End Sub


Private Sub flxORDER_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If ActRow = 0 Then Exit Sub
    If DateValue(flxORDER.TextMatrix(ActRow, 1)) <> Date Then Exit Sub
    If InStr(UCase$(flxORDER.TextMatrix(ActRow, 5)), "CANCELLED") > 0 Or Val(flxORDER.TextMatrix(ActRow, 6)) = 0 Then Exit Sub
    If Button = vbRightButton Then
       Call PopupMenu(mnuSalRet)
    End If
End Sub

Private Sub Form_Load()
    
    If Day(Date) <= 7 Then
        dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
    Else
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    End If
    dTo.Value = Date
    cmdNext.Tag = "2"
    Call RefreshRecord(2)
    ActRow = 0
    MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Return Summary"
    MotoRep.tbMain.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsSPord.State = adStateOpen Then rsSPord.Close
End Sub

Private Sub mnuCancel_Click()
    Dim rsCDtls As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim Ans As Integer, MyReason As String, Qry As String, MyDetails As String, MyAcc As String
    Dim MyUnit As String, MyDesc As String, MyInv As String, MyProc As String
    Dim i As Integer, CustID As Integer, InvBal As Single, TtlAmtPd As Currency, MyCredit As Currency, MyDebit As Currency
    
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=29));"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    'Display Details in flxCancel
    Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [RECEIVE]*[PRICE] AS AMOUNT, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE (((SPITR.REFERENCE)='RET#" & flxORDER.TextMatrix(ActRow, 2) & "') AND ((SPITR.TRANSTYPE)='SALES RETURN'));"
    With rsCDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancel.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            MyDetails = ""
            Do While Not .EOF
                '0-QTY,  1-PARTNO,  2-DESCRIPTION, 3-PRICE, 4-AMOUNT, 5-SPID
                flxCancel.AddItem !RECEIVE & vbTab & !PARTNUMBER & vbTab & !Description & vbTab & !Price & vbTab & !Amount & vbTab & !SPAREPARTID
                If InStr(UCase$(!Description), "4T OIL") > 0 Then
                    MyUnit = IIf(!RECEIVE > 1, "QRTS", "QRT")
                Else
                    MyUnit = IIf(!RECEIVE > 1, "PCS", "PC")
                End If
                If Len(!Description) < 30 Then
                    MyDesc = !Description & Space$(30 - Len(!Description))
                Else
                    MyDesc = Left$(!Description, 30)
                End If
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !RECEIVE & vbTab & MyUnit & vbTab & !PARTNUMBER & Space$(20 - Len(!PARTNUMBER)) & vbTab & MyDesc & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
                .MoveNext
            Loop
        End If
        .Close
    End With
    Ans = MsgBox("Do you want to Cancel Sales Return?" & vbCrLf & vbCrLf & _
        "Sales Return No.: " & flxORDER.TextMatrix(ActRow, 2) & vbTab & vbTab & vbTab & vbTab & "Sales Return Date : " & flxORDER.TextMatrix(ActRow, 1) & vbCrLf & vbCrLf & _
        "QTY." & vbTab & "UNIT" & vbTab & "PART NUMBER" & vbTab & vbTab & "DESCRIPTION" & vbTab & vbTab & vbTab & "PRICE" & vbTab & "AMOUNT" & vbCrLf & MyDetails, vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Sales Return:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Retrace transaction made
    'Update Sales Return Remarks to CANCELLED
    LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = 'CANCELLED', SRETURNSP.TOTALAMT = 0 WHERE (((SRETURNSP.RETNO)='" & flxORDER.TextMatrix(ActRow, 2) & "'));"
    
    With flxCancel
        For i = 1 To .Rows - 1
            'Cancel SPTIR
            LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.RECEIVE = 0 " & _
                           "WHERE (((SPITR.REFERENCE)='RET#" & flxORDER.TextMatrix(ActRow, 2) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'Update Qty-On Hand
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]-" & Val(.TextMatrix(i, 0)) & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
        Next i
    End With
    
    Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxORDER.TextMatrix(ActRow, 3) & "'));"
    If CtrlNumExist(Qry) Then   'Update Source Invoice field remarks to Null
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = Null WHERE (((SPINVOICE.SPINVNO)='" & flxORDER.TextMatrix(ActRow, 3) & "'));"
    End If
    Qry = "SELECT PAYMENTS.INVOICENO From PAYMENTS WHERE (((PAYMENTS.INVOICENO)='" & flxORDER.TextMatrix(ActRow, 3) & "'));"
    If CtrlNumExist(Qry) Then   'If Cash Update Payment's remarks to null
        LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = Null WHERE (((PAYMENTS.INVOICENO)='" & flxORDER.TextMatrix(ActRow, 3) & "'));"
    End If
    'Update Ledger and On-Account Tables if Credit
    CustID = GetFieldValue("SELECT SPINVOICE.CUSTOMERID, SPINVOICE.SPINVNO FROM SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxORDER.TextMatrix(ActRow, 3) & "'));")
    Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.REFNO)='" & flxORDER.TextMatrix(ActRow, 2) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & "));"
    If CtrlNumExist(Qry) Then
        'Cancel Sales Return in Ledger
        LYRcon.Execute "UPDATE LEDGER SET LEDGER.REFTYPE = 'CANCELLED', LEDGER.CUSTOMERID = 0 WHERE (((LEDGER.REFNO)='" & flxORDER.TextMatrix(ActRow, 2) & "'));"
        'Update OS Bal in Ledger
        InvBal = Val(GetFieldValue("SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxORDER.TextMatrix(ActRow, 3) & "'));"))
        TtlAmtPd = 0
        Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.SPINVNO)='" & flxORDER.TextMatrix(ActRow, 3) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ")) ORDER BY LEDGER.SEQNO;"
        With rsLedger   'Update Ledger
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    MyCredit = IIf(IsNull(!Credit), 0, !Credit)
                    TtlAmtPd = TtlAmtPd + MyCredit
                    MyDebit = IIf(IsNull(!Debit), 0, !Debit)
                    If MyCredit > 0 Then
                        InvBal = InvBal - MyCredit
                    ElseIf MyDebit > 0 Then
                        InvBal = InvBal + MyDebit
                    End If
                    !OSbal = InvBal
                    .MoveNext
                Loop
            End If
            .Close
        End With
        'Update Onaccount table
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & TtlAmtPd & ", ONACCTERMS.BALANCE = " & InvBal & ", ONACCTERMS.STATUS = 'OPEN', ONACCTERMS.LASTUPDATE = #" & Now & "# " & _
                       "WHERE (((ONACCTERMS.SPINVNO)='" & flxORDER.TextMatrix(ActRow, 3) & "'));"
        'Update Customer Outstanding Balance
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = [CUSTOMER].[OSBAL]+" & PureNum(flxORDER.TextMatrix(ActRow, 6)) & " WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & "));"
    End If
    'Update Audit Trans
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'SALES RETURN', '" & flxORDER.TextMatrix(ActRow, 2) & "', #" & flxORDER.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    'Update Grid
    With flxORDER
        .TextMatrix(ActRow, 7) = "CANCELLED"
        .TextMatrix(ActRow, 6) = "0.00"
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
        .Col = 2
    End With
    MsgBox "Sales Return Cancelled!", vbInformation + vbOKOnly, "Cancel"
                    
End Sub

Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxORDER
        .FormatString = "   |^RETURN DATE|^RETURN NO.|^INVOICE NO.|<CUSTOMER|<REASON |>TOTAL PRICE|<REMARKS |^PRINTED"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.1
        .ColWidth(2) = (.Width - 300) * 0.1
        .ColWidth(3) = (.Width - 300) * 0.1
        .ColWidth(4) = (.Width - 300) * 0.2
        .ColWidth(5) = (.Width - 300) * 0.2
        .ColWidth(6) = (.Width - 300) * 0.1
        .ColWidth(7) = (.Width - 300) * 0.2
        .ColWidth(8) = (.Width - 300) * 0.08
        .ColWidth(9) = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
            .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
            If Val(.TextMatrix(i, 8)) = 0 Then  'PRINT COPY
                .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
            End If
            If .TextMatrix(i, 7) = "CANCELLED" Then 'REMARKS
                .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
                .Col = 2
            End If
        Next i
        .ZOrder
    End With
    
    cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
    
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

