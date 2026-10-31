VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form DmCmSumm 
   Caption         =   "DMCM [A/R] Summary"
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
         Image           =   "DMCMsumm.frx":0000
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
         Format          =   49283073
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
         Image           =   "DMCMsumm.frx":059A
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
         Image           =   "DMCMsumm.frx":0B34
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
         Format          =   49283073
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
         Image           =   "DMCMsumm.frx":10CE
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
      Image           =   "DMCMsumm.frx":1668
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
      Image           =   "DMCMsumm.frx":1C02
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
      Image           =   "DMCMsumm.frx":219C
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxInvCancel 
      Height          =   735
      Left            =   1920
      TabIndex        =   17
      Top             =   8640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "REFERENCE NO.|^REF. TYPE"
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
      Caption         =   "DMCM ADJUSTMENT SUMMARY"
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
      Caption         =   "DMCM ADJUSTMENT SUMMARY"
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
Attribute VB_Name = "DmCmSumm"
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
        Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
              "Where (((ADJDMCM.ADJDATE) = #" & dFrom.Value & "#)) ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
    Else
        Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
              "Where (((ADJDMCM.ADJDATE) >= #" & dFrom.Value & "# And (ADJDMCM.ADJDATE) <= #" & dTo.Value & "#)) ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
    End If
    If rsAdj.State = adStateOpen Then rsAdj.Close
    rsAdj.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsAdj.RecordCount = 0 Then
        flxAdj.Clear
        flxAdj.Rows = 2
        flxAdj.Cols = 11
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
            For j = fCol To 11
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 11:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 11, i + 1, i)
                    fCol = IIf(j = 11, 1, j + 1)
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
        Case 1  'Show DMCM
            With flxAdj
                If ActRow = 0 Or ActRow > .Rows - 1 Then
                    MsgBox "Please select a record to view!", vbInformation + vbOKOnly, "DMCM Summary"
                    Exit Sub
                End If
                Call DMCMadj.ShowRecord(flxAdj.TextMatrix(ActRow, 2))
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
    MotoRep.sBar.Panels(6).Text = "Inventory Adjustment Summary"
    MotoRep.tbMain.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsAdj.State = adStateOpen Then rsAdj.Close
End Sub


Private Sub mnuCancel_Click()
    Dim Ans As Integer, Bal As Currency, AmtAdj As Currency
    Dim Qry As String, MyTrans As String, RefType As String, AdjType As String, MyStat As String, JOrem As String
    Dim MyReason As String, MyAcc As String
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=54));"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel Adjustment?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of DMCM:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Adjustment failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
        
    MyTrans = flxAdj.TextMatrix(ActRow, 4)
    If Val(Replace(flxAdj.TextMatrix(ActRow, 7), ",", "")) > 0 Then
        AmtAdj = Val(Replace(flxAdj.TextMatrix(ActRow, 7), ",", ""))
        AdjType = "CREDIT"
    Else
        AmtAdj = Val(Replace(flxAdj.TextMatrix(ActRow, 6), ",", ""))
        AdjType = "DEBIT"
    End If
    Bal = Val(Replace(flxAdj.TextMatrix(ActRow, 8), ",", ""))
    'UPDATE AdjDMCM
    LYRcon.Execute "UPDATE ADJDMCM SET ADJDMCM.REASON = 'CANCELLED', ADJDMCM.CUSTOMERID = 0 WHERE (((ADJDMCM.ADJNO)='" & flxAdj.TextMatrix(ActRow, 2) & "'));"
    
    If MyTrans = "INVOICE" Then
        Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxAdj.TextMatrix(ActRow, 4) & "'));"
        RefType = GetFieldValue(Qry)
        If RefType = "CREDIT" Then  'INVOICE TYPE = CREDIT
            Call UpdateLedger(flxAdj.TextMatrix(ActRow, 2))
            LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = Null WHERE (((SPINVOICE.SPINVNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
            'update onaccount
            If AdjType = "CREDIT" Then  'DMCM ADJ TYPE = CREDIT
                MyStat = IIf(Bal + AmtAdj > 0, "OPEN", "CLOSED")
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal + AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.SPINVNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
            Else
                MyStat = IIf(Bal - AmtAdj > 0, "OPEN", "CLOSED")
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal - AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.SPINVNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
            End If
        ElseIf RefType = "CASH" Then
            If AdjType = "CREDIT" Then
                LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.AMOUNT = " & Bal + AmtAdj & ", SPINVOICE.REMARKS = Null " & _
                               "WHERE (((SPINVOICE.SPINVNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
            ElseIf AdjType = "DEBIT" Then
                LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.AMOUNT = " & Bal - AmtAdj & ", SPINVOICE.REMARKS = Null " & _
                               "WHERE (((SPINVOICE.SPINVNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
            End If
        End If
    ElseIf MyTrans = "JOB ORDER" Then
        Qry = "SELECT SRVCJOBORDER.PAYTYPE From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
        RefType = GetFieldValue(Qry)
        Qry = "SELECT SRVCJOBORDER.REMARKS From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
        JOrem = GetFieldValue(Qry)
        JOrem = Replace(JOrem, "DMCM ADJ# " & flxAdj.TextMatrix(ActRow, 2), "")
        If RefType = "ON ACCOUNT" Then  'JOB ORDER PAY TYPE = ON ACCOUNT
            Call UpdateLedger(flxAdj.TextMatrix(ActRow, 2))
            LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                       "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
            'update onaccount
            If AdjType = "CREDIT" Then  'DMCM ADJ TYPE = CREDIT
                MyStat = IIf(Bal + AmtAdj > 0, "OPEN", "CLOSED")
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal + AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.JOBORDERNO)='" & flxAdj.TextMatrix(ActRow, 5) & "') AND ((ONACCTERMS.SPINVNO) Is Null));"
            Else
                MyStat = IIf(Bal - AmtAdj > 0, "OPEN", "CLOSED")
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal - AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.JOBORDERNO)='" & flxAdj.TextMatrix(ActRow, 5) & "') AND ((ONACCTERMS.SPINVNO) Is Null));"
            End If
        ElseIf RefType = "CASH" Then
            If AdjType = "CREDIT" Then
                LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Bal + AmtAdj & ", SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                               "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
            ElseIf AdjType = "DEBIT" Then
                LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Bal - AmtAdj & ", SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                               "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxAdj.TextMatrix(ActRow, 5) & "'));"
            End If
        End If
    End If
   'Update Grid
    flxAdj.TextMatrix(ActRow, 10) = "CANCELLED"
    With flxAdj     'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 11: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
    End With
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                   "VALUES (#" & Date & "#, 'CANCEL', 'DMCM ADJ.', '" & flxAdj.TextMatrix(ActRow, 2) & "', #" & flxAdj.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    
    MsgBox "Adjustment Cancelled!", vbInformation + vbOKOnly
End Sub

Private Sub UpdateLedger(MyRefNo As String)
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, CustID As Integer
    Dim MyCredit As Currency, MyDebit As Currency, InvBal As Currency
    
    Qry = "SELECT LEDGER.* FROM LEDGER WHERE (((LEDGER.REFNO)='" & MyRefNo & "') AND ((LEDGER.REFTYPE)='DMCM'));"
    With rsLedger   'Update Ledger
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxInvCancel.Rows = 1
            CustID = !CUSTOMERID
            Do While Not .EOF
                MyJOid = IIf(IsNull(!JobID), 0, !JobID)
                flxInvCancel.AddItem !SPINVNO & vbTab & IIf(MyJOid > 0, "JOB", "INV")
                .MoveNext
            Loop
        End If
        .Close
    End With
    For i = 1 To flxInvCancel.Rows - 1
        If flxInvCancel.TextMatrix(i, 1) = "INV" Then
            Qry = "SELECT LEDGER.* FROM LEDGER WHERE (((LEDGER.SPINVNO)='" & flxInvCancel.TextMatrix(i, 0) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ")) ORDER BY LEDGER.SEQNO;"
            InvBal = GetInvAmt(flxInvCancel.TextMatrix(i, 0))
        Else
            Qry = "SELECT ONACCTERMS.ACCTAMOUNT FROM ONACCTERMS WHERE (((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.JOBORDERNO)='" & flxInvCancel.TextMatrix(i, 0) & "'));"
            InvBal = GetFieldValue(Qry)
            Qry = "SELECT LEDGER.* From Ledger WHERE (((LEDGER.SPINVNO)='" & flxInvCancel.TextMatrix(i, 0) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ") AND ((LEDGER.JOBID)>0)) ORDER BY LEDGER.SEQNO;"
        End If
        With rsLedger   'Update Ledger
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If !RefNo <> MyRefNo And !RefType <> "DMCM" Then
                        MyCredit = IIf(IsNull(!Credit), 0, !Credit)
                        TtlAmtPd = TtlAmtPd + MyCredit
                        MyDebit = IIf(IsNull(!Debit), 0, !Debit)
                        If MyCredit > 0 Then
                            InvBal = InvBal - MyCredit
                        ElseIf MyDebit > 0 Then
                            InvBal = InvBal + MyDebit
                        End If
                        !OSbal = InvBal
                    End If
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Next i
    'Update Ledger's RefType to CANCELLED and CustomerID=0 where RefNo=MyOR
    LYRcon.Execute "UPDATE Ledger SET Ledger.REFTYPE = 'CANCELLED', Ledger.CUSTOMERID = 0 WHERE (((Ledger.REFTYPE)='DMCM') AND ((Ledger.REFNO)='" & MyRefNo & "'));"
End Sub

Function GetInvAmt(InvNo As String) As Single
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & InvNo & "'));"
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetInvAmt = !Amount
        Else
            GetInvAmt = 0
        End If
        .Close
    End With
    
End Function


Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxAdj         '       1       2           3       4         5         6       7        8       9        10
        .FormatString = "   |^ DATE|^ADJ. NO.|<CUSTOMER|<REFERENCE|^REF. NO.|>DEBIT |>CREDIT|>BALANCE|<STATUS|<REASON|<PRN COPY"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.2
        .ColWidth(4) = (.Width - 300) * 0.12
        .ColWidth(5) = (.Width - 300) * 0.12
        .ColWidth(6) = (.Width - 300) * 0.12
        .ColWidth(7) = (.Width - 300) * 0.12
        .ColWidth(8) = (.Width - 300) * 0.12
        .ColWidth(9) = (.Width - 300) * 0.12
        .ColWidth(10) = (.Width - 300) * 0.2
        .ColWidth(10) = (.Width - 300) * 0.12
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
            .TextMatrix(i, 8) = Format$(.TextMatrix(i, 8), "#,##0.00")
            .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
            .TextMatrix(i, 7) = Format$(.TextMatrix(i, 7), "#,##0.00")
            If Val(.TextMatrix(i, 11)) = 0 Then
                '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
            End If
            If .TextMatrix(i, 10) = "CANCELLED" Then
                '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
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

