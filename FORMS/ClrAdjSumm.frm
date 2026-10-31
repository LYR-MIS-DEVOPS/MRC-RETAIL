VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form ClrAdjSumm 
   Caption         =   "Clerical Adjustment Summary"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAdj 
      Height          =   7095
      Left            =   120
      TabIndex        =   13
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
         Image           =   "ClrAdjSumm.frx":0000
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
         Format          =   49217537
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   12
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
         Image           =   "ClrAdjSumm.frx":059A
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   14
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
         Image           =   "ClrAdjSumm.frx":0B34
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   15
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49217537
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   16
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
         Image           =   "ClrAdjSumm.frx":10CE
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
         TabIndex        =   11
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Adjustment"
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
      Image           =   "ClrAdjSumm.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   9
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
      Image           =   "ClrAdjSumm.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   10
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
      Image           =   "ClrAdjSumm.frx":219C
      cBack           =   -2147483633
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
      Caption         =   "CLERICAL ADJUSTMENT SUMMARY"
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
      Caption         =   "CLERICAL ADJUSTMENT SUMMARY"
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
   Begin VB.Menu mnuADJ 
      Caption         =   "Clerical Adj summary"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Adjustment"
      End
   End
End
Attribute VB_Name = "ClrAdjSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAdj As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Dim Qry As String

Private Sub RefreshRecord(dCode As Integer)
    Dim i As Integer
    
    If dCode = 1 Then
        Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
              "From ADJCLERICAL Where (((ADJCLERICAL.ADJDATE) = #" & dFrom.Value & "#)) ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
    Else
        Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
              "From ADJCLERICAL Where (((ADJCLERICAL.ADJDATE) >= #" & dFrom.Value & "# And (ADJCLERICAL.ADJDATE) <= #" & dTo.Value & "#)) ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
    End If
    If rsAdj.State = adStateOpen Then rsAdj.Close
    rsAdj.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsAdj.RecordCount = 0 Then
        flxAdj.Clear
        flxAdj.Rows = 2
        flxAdj.Cols = 10
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxAdj.DataSource = rsAdj
    ActRow = 1
    flxAdj.Visible = True
    Call RXResizer1_ResizeComplete
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxAdj
        For i = fRow To .Rows - 1
            For j = fCol To 10
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 10:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 10, i + 1, i)
                    fCol = IIf(j = 10, 1, j + 1)
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
    Dim dCov As String, Qry As String
    
    Select Case Index
        Case 1  'Show Inventory Adjustment
            With flxAdj
                If ActRow = 0 Or ActRow > .Rows - 1 Then
                    MsgBox "Please select a record to view!", vbInformation + vbOKOnly, "Confirm"
                    Exit Sub
                End If
                Call ClericalADJ.DisplayClrADJ(flxAdj.TextMatrix(ActRow, 2))
                ClericalADJ.Show vbModal
            End With
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

Private Sub flxAdj_EnterCell()
    ActRow = flxAdj.Row
End Sub

Private Sub flxAdj_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If DateValue(flxAdj.TextMatrix(ActRow, 1)) <> Date Then Exit Sub
    If InStr(UCase$(flxAdj.TextMatrix(ActRow, 9)), "CANCELLED") > 0 Or Val(flxAdj.TextMatrix(ActRow, 11)) = 0 Then Exit Sub
    If Button = vbRightButton Then
       Call PopupMenu(mnuADJ)
    End If
End Sub

Private Sub Form_Load()
    
    Call cmdNext_Click
    dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dTo.Value = DateAdd("d", -1, DateSerial(Year(Date), Month(Date) + 1, 1))
    Call RefreshRecord(2)
    ActRow = 1
    MotoRep.sBar.Panels(6).Text = "Clerical Adjustment Summary"
    MotoRep.tbMain.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsAdj.State = adStateOpen Then rsAdj.Close
End Sub

Private Sub mnuCancel_Click()
    Dim Ans As Integer, MyReason As String, Qry As String, MyRefType As String, MyAcc As String
    Dim ReqID As Long

    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=56));"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If

    Ans = MsgBox("Do you want to Cancel Adjustment?" & vbCrLf & vbCrLf & _
        "Adj. Date : " & flxAdj.TextMatrix(ActRow, 1) & vbCrLf & "Adjustment No: " & flxAdj.TextMatrix(ActRow, 2) & vbCrLf & _
        "Transaction: " & flxAdj.TextMatrix(ActRow, 3), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Adjustment:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Adjustment failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    '...Retrace transactions base on Item Adjusted
    With flxAdj
        If .TextMatrix(ActRow, 3) = "PART NUMBER" Then
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.PARTNUMBER = '" & .TextMatrix(ActRow, 5) & "' " & _
                   "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(ActRow, 6) & "'));"
            LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.PARTNUMBER = '" & .TextMatrix(ActRow, 5) & "' " & _
                           "WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(ActRow, 6) & "'));"
        End If
        If .TextMatrix(ActRow, 3) = "INVOICE" And .TextMatrix(ActRow, 4) = "INVOICE NUMBER" Then
            Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 6) & "'));"
            If CtrlNumExist(Qry) Then
                MsgBox "Unable to Cancel Transaction because Invoice No. " & .TextMatrix(ActRow, 6) & " was already used.", vbExclamation + vbOKOnly, "Cancel Failed"
                Exit Sub
            End If
            'Cancel Invoice Number
            Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 7) & "'));"
            MyRefType = GetFieldValue(Qry)
            eMsg = "Line 170: Updating Invoice Number"
            LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.SPINVNO = '" & .TextMatrix(ActRow, 6) & "', SPINVOICE.TRANSNO = 'SI#-" & .TextMatrix(ActRow, 6) & "' " & _
                           "WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 7) & "'));"
            eMsg = "Line 173: Updating Reference of Transaction Register"
            LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = 'SI#-" & .TextMatrix(ActRow, 6) & "' " & _
                           "WHERE (((SPITR.REFERENCE)='SI#-" & .TextMatrix(ActRow, 7) & "'));"
            If MyRefType = "CASH" Then
                eMsg = "Line 177: Updating Invoice Number on table Payments"
                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.INVOICENO = '" & .TextMatrix(ActRow, 6) & "' WHERE (((PAYMENTS.INVOICENO)='" & .TextMatrix(ActRow, 7) & "'));"
            Else
                eMsg = "Line 180: Updating Invoice Number on table On Account Terms"
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.SPINVNO = '" & .TextMatrix(ActRow, 6) & "' WHERE (((ONACCTERMS.SPINVNO)='" & .TextMatrix(ActRow, 7) & "'));"
                Qry = "SELECT LEDGER.SPINVNO From Ledger WHERE (((LEDGER.SPINVNO)='" & .TextMatrix(ActRow, 7) & "') AND ((LEDGER.JOBID)=0));"
                If CtrlNumExist(Qry) Then
                    eMsg = "Line 184: Updating Invoice Number on table Ledger"
                    LYRcon.Execute "UPDATE LEDGER SET LEDGER.SPINVNO = '" & .TextMatrix(ActRow, 6) & "' " & _
                            "WHERE (((LEDGER.SPINVNO)='" & .TextMatrix(ActRow, 7) & "') AND ((LEDGER.JOBID)=0));"
                End If
            End If
            Qry = "SELECT SPINVOICE.REQID From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 6) & "'));"
            ReqID = Val(GetFieldValue(Qry))
            If ReqID > 0 Then
                eMsg = "Line 199: Updating Invoice Number on table Requisition"
                LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = '" & .TextMatrix(ActRow, 6) & "' WHERE (((REQUISITION.REQID)=" & ReqID & "));"
            End If
        End If
        If .TextMatrix(ActRow, 3) = "INVOICE" And .TextMatrix(ActRow, 4) = "INVOICE DATE" Then
            Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 5) & "'));"
            If Not CtrlNumExist(Qry) Then
                MsgBox "Unable to Cancel Transaction because Invoice No. " & .TextMatrix(ActRow, 6) & " does not exist", vbExclamation + vbOKOnly, "Cancel Failed"
                Exit Sub
            End If
            Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 5) & "'));"
            MyRefType = GetFieldValue(Qry)
            eMsg = "Line 200: Updating Invoice Date"
            LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.INVDATE = #" & .TextMatrix(ActRow, 6) & "# WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 5) & "'));"
            eMsg = "Line 202: Updating Date on Transaction Register"
            LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSDATE = #" & .TextMatrix(ActRow, 6) & "# WHERE (((SPITR.REFERENCE)='SI#-" & .TextMatrix(ActRow, 5) & "'));"
            If MyRefType = "CASH" Then
                eMsg = "Line 205: Updating Invoice Date on table Payments"
                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.PAYDATE= #" & .TextMatrix(ActRow, 6) & "# WHERE (((PAYMENTS.INVOICENO)='" & .TextMatrix(ActRow, 5) & "'));"
            Else
                eMsg = "Line 208: Updating Invoice Date on table On Account Terms"
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TRANSDATE = #" & .TextMatrix(ActRow, 6) & "# WHERE (((ONACCTERMS.SPINVNO)='" & .TextMatrix(ActRow, 5) & "'));"
            End If
        End If
        LYRcon.Execute "UPDATE ADJCLERICAL SET ADJCLERICAL.REMARKS = 'CANCELLED' WHERE (((ADJCLERICAL.ADJNO)='" & .TextMatrix(ActRow, 2) & "'));"
    
        'Update Audit Trans
        LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                 "VALUES (#" & Date & "#, 'CANCEL', 'CLERICAL ADJ.', '" & .TextMatrix(ActRow, 2) & "', #" & .TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
        'Update Grid
         .TextMatrix(ActRow, 9) = "CANCELLED"
        'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 11: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
        MsgBox .TextMatrix(ActRow, 3) & " Adjustment Cancelled!", vbInformation + vbOKOnly, "Cancel"
    End With

End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxAdj             '   1       2          3          4       5            6        7        8         9          10
        .FormatString = "   |^ DATE|^ADJ. NO.|<TRANSACTION|<FIELD|<REF. NO.|<OLD VALUE |<NEW VALUE|<REASON |<REMARKS|<ENCODER|<PRINT COPY"
        .ColWidth(0) = 300
        For i = 1 To 10
            .ColWidth(i) = (.Width - 300) * 0.12
        Next i
        .ColWidth(11) = (.Width - 300) * 0.08
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
            If Val(.TextMatrix(i, 11)) = 0 Then
                .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
            End If
            If .TextMatrix(i, 9) = "CANCELLED" Then
                .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
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

