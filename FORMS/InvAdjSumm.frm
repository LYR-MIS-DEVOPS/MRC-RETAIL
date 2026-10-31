VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form InvAdjSumm 
   Caption         =   "Inventory Adjustment Summary"
   ClientHeight    =   8625
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxCancel 
      Height          =   1695
      Left            =   1560
      TabIndex        =   17
      Top             =   8640
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      FormatString    =   "Part Number|<Adj. Type|< IN      |< OUT   |<Source | SPID"
   End
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
         Image           =   "InvAdjSumm.frx":0000
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
         Format          =   51511297
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
         Image           =   "InvAdjSumm.frx":059A
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
         Image           =   "InvAdjSumm.frx":0B34
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
         Format          =   51511297
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
         Image           =   "InvAdjSumm.frx":10CE
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
      Image           =   "InvAdjSumm.frx":1668
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
      Image           =   "InvAdjSumm.frx":1C02
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
      Image           =   "InvAdjSumm.frx":219C
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
      m_MDI_Height    =   9135
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTORY ADJUSTMENT SUMMARY"
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
      Caption         =   "INVENTORY ADJUSTMENT SUMMARY"
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
      Caption         =   "ADJ summary"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Adjustment"
      End
   End
End
Attribute VB_Name = "InvAdjSumm"
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
        Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
              "From InvntryADJ Where (((InvntryADJ.ADJDATE) = #" & dFrom.Value & "#)) ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
    Else
        Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
              "From InvntryADJ Where (((InvntryADJ.ADJDATE) >= #" & dFrom.Value & "# And (InvntryADJ.ADJDATE) <= #" & dTo.Value & "#)) ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
    End If
    If rsAdj.State = adStateOpen Then rsAdj.Close
    rsAdj.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsAdj.RecordCount = 0 Then
        flxAdj.Clear
        flxAdj.Rows = 2
        flxAdj.Cols = 8
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
            For j = fCol To 8
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 10:  .RowSel = i
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
    Dim dCov As String, Qry As String
    
    Select Case Index
        Case 1  'Show Inventory Adjustment
            With flxAdj
                If ActRow = 0 Or ActRow > .Rows - 1 Then
                    MsgBox "Please select a record to view!", vbInformation + vbOKOnly, "Confirm"
                    Exit Sub
                End If
                Call InvntryADJ.DisplayInvAdj(flxAdj.TextMatrix(ActRow, 2))
            End With
            Unload Me
        Case 2  'Print
            'dCov = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(cmdNext.Tag <> "1", " to " & Format(dFrom.Value, "Mmm. dd, yyyy"), "")
            'If Val(cmdNext.Tag) = 1 Then
            '    Qry = "SELECT SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS " & _
            '          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATERECEIVED) = #" & dFrom.Value & "#)) ORDER BY SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOBORDERNO;"
            'Else
            '    Qry = "SELECT SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS " & _
            '          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATERECEIVED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATERECEIVED) <= #" & dTo.Value & "#)) ORDER BY SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOBORDERNO;"
            'End If
            'PRNmc.QryPrn = Qry
            'PRNmc.ParamValue1 = dCov
            'PRNmc.ParamValue2 = GetUserName
            'PRNmc.MySource = "JOSUMMARY"
            'PRNmc.Show
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
    If InStr(UCase$(flxAdj.TextMatrix(ActRow, 5)), "CANCELLED") > 0 Or Val(flxAdj.TextMatrix(ActRow, 8)) = 0 Then Exit Sub
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
    MotoRep.sBar.Panels(6).Text = "Inventory Adjustment Summary"
    MotoRep.tbMain.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsAdj.State = adStateOpen Then rsAdj.Close
End Sub

Private Sub mnuCancel_Click()
    Dim rsCDtls As New ADODB.Recordset
    Dim Ans As Integer, MyReason As String, Qry As String, MyDetails As String, MyAcc As String
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=55));"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    'Display Details in flxCancel
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPITR.REMARKS, SPITR.RECEIVE, SPITR.OUT, SPITR.SOURCE, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE (((SPITR.REFERENCE)='ADJ#" & flxAdj.TextMatrix(ActRow, 2) & "') AND ((SPITR.TRANSTYPE)='INVENTORY ADJ.'));"
    With rsCDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancel.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            MyDetails = ""
            Do While Not .EOF
                '0-Part Number,  1-Adj. Type,  2-IN, 3-OUT, 4-Source, 5-SPID
                flxCancel.AddItem !PARTNUMBER & vbTab & IIf(IsNull(!REMARKS), "", !REMARKS) & vbTab & !RECEIVE & vbTab & !OUT & vbTab & IIf(IsNull(!Source), "", !Source) & vbTab & !SPAREPARTID
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !PARTNUMBER & vbTab & flxCancel.TextMatrix(flxCancel.Rows - 1, 1) & Space$(2) & vbTab & IIf(InStr(UCase$(flxCancel.TextMatrix(flxCancel.Rows - 1, 1)), "DEBIT") > 0, !RECEIVE, !OUT) & vbTab & flxCancel.TextMatrix(flxCancel.Rows - 1, 4)
                .MoveNext
            Loop
        End If
        .Close
    End With
    Ans = MsgBox("Do you want to Cancel Adjustment?" & vbCrLf & vbCrLf & _
        "Adjustment No: " & flxAdj.TextMatrix(ActRow, 2) & vbTab & "Adjustment Date : " & flxAdj.TextMatrix(ActRow, 1) & vbCrLf & vbCrLf & _
        "PART NUMBER" & vbTab & "ADJ. TYPE" & vbTab & "QTY." & vbTab & "SOURCE" & vbCrLf & MyDetails, vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Adjustment:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Adjustment failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Retrace transaction made
    LYRcon.Execute "UPDATE INVNTRYADJ SET INVNTRYADJ.REMARKS = 'CANCELLED' WHERE (((INVNTRYADJ.ADJNO)='" & flxAdj.TextMatrix(ActRow, 2) & "'));"
    Qry = "SELECT INVNTRYADJ.APPLYCHARGES, INVNTRYADJ.ADJNO From InvntryADJ WHERE (((INVNTRYADJ.ADJNO)='" & flxAdj.TextMatrix(ActRow, 2) & "'));"
    If GetFieldValue(Qry) = True Then   'Update Apply Charges if true
        LYRcon.Execute "UPDATE INVADJCHARGES SET INVADJCHARGES.EMPID = 0 WHERE (((INVADJCHARGES.ADJNO)='" & flxAdj.TextMatrix(ActRow, 2) & "'));"
    End If
    'Update Details
    With flxCancel
        For i = 1 To .Rows - 1
            If InStr(UCase$(.TextMatrix(i, 1)), "DEBIT") > 0 Then
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]-" & Val(.TextMatrix(i, 2)) & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
                LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.RECEIVE = 0 " & _
                               "WHERE (((SPITR.REFERENCE)='ADJ#" & flxAdj.TextMatrix(ActRow, 2) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            Else
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+" & Val(.TextMatrix(i, 3)) & _
                               " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
                LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.OUT = 0 " & _
                               "WHERE (((SPITR.REFERENCE)='ADJ#" & flxAdj.TextMatrix(ActRow, 2) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            End If
        Next i
    End With
    
    'Update Audit Trans
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'INVENTORY ADJ.', '" & flxAdj.TextMatrix(ActRow, 2) & "', #" & flxAdj.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
   'Update Grid
    flxAdj.TextMatrix(ActRow, 5) = "CANCELLED"
    With flxAdj     'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
    End With
    MsgBox "Adjustment Cancelled!", vbInformation + vbOKOnly

End Sub

Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxAdj         '       1       2           3       4       5           6           7           8
        .FormatString = "   |^ DATE|^ADJ. NO.|<REFERENCE|<REASON|<REMARKS |<PREPARED BY|<APPROVED BY|<PRINT COPY"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.12
        .ColWidth(4) = (.Width - 300) * 0.2
        .ColWidth(5) = (.Width - 300) * 0.2
        .ColWidth(6) = (.Width - 300) * 0.2
        .ColWidth(7) = (.Width - 300) * 0.2
        .ColWidth(8) = (.Width - 300) * 0.15
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
            If Val(.TextMatrix(i, 8)) = 0 Then
                '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
            End If
            If .TextMatrix(i, 5) = "CANCELLED" Then
                '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
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

