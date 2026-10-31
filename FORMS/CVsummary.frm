VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form CVsummary 
   Caption         =   "Cash Voucher Summary"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxCV 
      Height          =   6975
      Left            =   0
      TabIndex        =   15
      Tag             =   "XY"
      Top             =   1200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   275
      BackColorBkg    =   -2147483633
      AllowUserResizing=   1
      FormatString    =   $"CVsummary.frx":0000
   End
   Begin CRVIEWERLibCtl.CRViewer vwrDmCm 
      Height          =   7695
      Left            =   0
      TabIndex        =   10
      Tag             =   "XY"
      Top             =   480
      Width           =   11415
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   400
      Width           =   11295
      Begin VB.TextBox txtYear 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3495
         TabIndex        =   5
         Text            =   "2007"
         Top             =   240
         Width           =   705
      End
      Begin VB.ComboBox cboMonth 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1935
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.UpDown updYear 
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2006
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtYear"
         BuddyDispid     =   196610
         OrigLeft        =   4095
         OrigTop         =   240
         OrigRight       =   4350
         OrigBottom      =   555
         Max             =   2015
         Min             =   2006
         Enabled         =   -1  'True
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "For the Month of"
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
         Height          =   255
         Index           =   2
         Left            =   -720
         TabIndex        =   8
         Top             =   285
         Width           =   2505
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "For the Month of"
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
         Left            =   -720
         TabIndex        =   9
         Top             =   315
         Width           =   2505
      End
      Begin VB.Label lblCovered 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE COVERED: DEC. 1 TO DEC. 31, 2006"
         Height          =   255
         Left            =   6960
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   8280
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   7800
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   8280
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   7800
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Voucher"
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
      Image           =   "CVsummary.frx":00D7
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Grid View"
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
      Image           =   "CVsummary.frx":0671
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "CVsummary.frx":0C0B
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "CVsummary.frx":11A5
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "CVsummary.frx":173F
      cBack           =   -2147483633
   End
   Begin VB.Label lblRow 
      Height          =   255
      Left            =   10320
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblP 
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   8280
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11625
      m_MDI_Height    =   8685
   End
   Begin VB.Menu mnuVoucher 
      Caption         =   "Cancel CV"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Voucher"
      End
   End
End
Attribute VB_Name = "CVsummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer, Qry As String
Dim ActRow As Integer, dFrm As Date, dTo As Date
Public ViewList As Boolean


Private Sub cboMonth_Click()
    Dim MyPart As Variant
    
    dFrom = DateSerial(txtYear.Text, cboMonth.ListIndex + 1, 1)
    dTo = DateAdd("d", -1, DateAdd("m", 1, dFrom))
    Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY From CASHVOUCHER " & _
          "Where (((CASHVOUCHER.CVDate) >= #" & dFrom & "# And (CASHVOUCHER.CVDate) <= #" & dTo & "#)) ORDER BY CASHVOUCHER.CVDATE DESC, CASHVOUCHER.VOUCHERNO;"

    If rsRPT.State = adStateOpen Then rsRPT.Close
    With rsRPT
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxCV.Rows = 1
            .MoveFirst
            Do While Not rsRPT.EOF
                MyPart = Split(!Particular, "|")
                flxCV.AddItem vbTab & Format$(!CVDate, "mm/dd/yyyy") & vbTab & !VOUCHERNO & vbTab & !PaidTo & vbTab & IIf(IsNull(!Position), "", !Position) & vbTab & MyPart(0) & " " & MyPart(1) & " " & MyPart(2) & vbTab & FormatNumber(!Amount, 2) & vbTab & !Remarks & vbTab & !PrintCopy
                .MoveNext
            Loop
        Else
            flxCV.Rows = 1
            flxCV.Rows = 2
        End If
    End With
    'Call FlxColumnResize
    MotoRep.sBar.Panels(6).Text = "Cash Voucher Summary"
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim MyQry As String
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
            Call cboMonth_Click
        Case 1  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
        Case 2  'view
            If ActRow = 0 Or flxCV.Rows = 2 Then ActRow = 1
            With flxCV
                'CashVoucher.CVtype = flxCV.TextMatrix(ActRow, 7)
                Call CashVoucher.ShowRecord(flxCV.TextMatrix(ActRow, 2))
                CashVoucher.Show vbModal
            End With
        Case 3  'Print Preview
            BtnIndx = 4
            tmrFloatRev.Enabled = True
            Call PrintPreView(True)
            Call DisplayReport
        Case 4  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
    End Select
End Sub

Private Sub flxCV_Click()
    If flxCV.Col = 0 And flxCV.Row > 0 Then
        ActRow = flxCV.Row
        Call flxcv_EnterCell
    End If
End Sub

Private Sub flxcv_EnterCell()
    ActRow = flxCV.Row
    lblRow.Caption = ActRow
    With flxCV
        .Row = ActRow: .Col = .Cols - 1
        .RowSel = ActRow: .ColSel = 0
        .LeftCol = 0
    End With
End Sub


Private Sub flxCV_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton And Val(flxCV.TextMatrix(ActRow, 8)) > 0 Then
        Call PopupMenu(mnuVoucher)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer

    
    For i = 1 To 12
        cboMonth.AddItem MonthName(i)
    Next i
    updYear.Value = Year(Date)
    txtYear.Text = Year(Date)
    cboMonth.Text = cboMonth.List(Month(Date) - 1)
    cboMonth.ListIndex = Month(Date) - 1
    dFrom = DateSerial(txtYear.Text, cboMonth.ListIndex + 1, 1)
    dTo = DateAdd("d", -1, DateAdd("m", 1, dFrom))
    lblCovered.Caption = "Date Covered : " & Format(dFrom, "MMM. dd") & " to " & Format(dTo, "MMM. dd, yyyy")
    
    If ViewList Then
        BtnIndx = 1
        tmrFloat.Enabled = True
        Call PrintPreView(False)
    Else
        BtnIndx = 4
        tmrFloatRev.Enabled = True
        Call PrintPreView(True)
        Call DisplayReport
    End If
End Sub

Private Sub DisplayReport()
    'Dim crARdmcm As New ARdmcm
    
    'dFrom = DateSerial(txtYear.Text, cboMonth.ListIndex + 1, 7)
    'dTo = DateAdd("d", -1, DateAdd("m", 1, dFrom))
    'lblCovered.Caption = "Date Covered : " & Format(dFrom, "MMM. dd, yyyy") & " to " & Format(dTo, "MMM. dd, yyyy")
    
    'Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY From CASHVOUCHER " & _
          "Where (((CASHVOUCHER.CVDate) >= #" & dFrom & "# And (CASHVOUCHER.CVDate) <= #" & dTo & "#)) ORDER BY CASHVOUCHER.CVDATE DESC, CASHVOUCHER.VOUCHERNO;"
    
    'If rsRPT.State = adStateOpen Then rsRPT.Close
    'rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    'crARdmcm.DiscardSavedData
    'crARdmcm.Database.SetDataSource rsRPT
    'crARdmcm.EnableParameterPrompting = False
    'crARdmcm.ParameterFields(1).ClearCurrentValueAndRange
    'crARdmcm.ParameterFields(1).AddCurrentValue lblCovered.Caption
    'vwrDmCm.ReportSource = crARdmcm
    'vwrDmCm.ViewReport
    MotoRep.sBar.Panels(6).Text = "Debit/Credit Memo Print Preview"
End Sub

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
        Frame1.Visible = False
        flxCV.Visible = False
        vwrDmCm.Visible = True
    Else
        vwrDmCm.Visible = False
        Frame1.Visible = True
        flxCV.Visible = True
    End If
End Sub

Private Sub FlxColumnResize()
    Dim i As Integer
    
    With flxCV
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 2500
        .ColAlignment(4) = 1
        For i = 5 To 11
            .ColWidth(i) = 1500
        Next i
        .ColWidth(12) = 3000
        .ColAlignment(12) = 1
        .ColWidth(13) = 1700
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 13)) = 0 Then
                .FillStyle = flexFillSingle
                .Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .Text = lblP.Caption
                .Col = 14: .Row = i
                .RowSel = i: .ColSel = 1:
                .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 150)
                .Col = 2
            End If
        Next
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
End Sub

Private Sub mnuCancel_Click()
    Dim Ans As Integer, MyReason As String, Qry As String, MyAcc As String
    Dim myAcctDesc As String * 50
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=60));"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If

    Ans = MsgBox("Do you want to Cancel Voucher?" & vbCrLf & vbCrLf & _
        "C.V. Date : " & flxCV.TextMatrix(ActRow, 1) & vbCrLf & _
        "Voucher No: " & flxCV.TextMatrix(ActRow, 2) & vbCrLf & _
        "Received By: " & flxCV.TextMatrix(ActRow, 3) & vbCrLf & _
        "Particular : " & flxCV.TextMatrix(ActRow, 5), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Voucher:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Voucher failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    MyReason = Replace(MyReason, "'", "`")
    'Update Voucher
    LYRcon.Execute "UPDATE CASHVOUCHER SET CASHVOUCHER.REMARKS = 'CANCELLED' WHERE (((CASHVOUCHER.VOUCHERNO)=" & flxCV.TextMatrix(ActRow, 2) & "));"
    'Add to Disbursement
    myAcctDesc = "CASH VOUCHER [" & flxCV.TextMatrix(ActRow, 5) & "]"
    LYRcon.Execute "INSERT INTO DISBURSEMENT ( CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, AMOUNT, STATUS, PRINTCOPY, UPDATEDBY, LASTUPDATE ) " & _
                   "VALUES (#" & flxCV.TextMatrix(ActRow, 1) & "#, " & flxCV.TextMatrix(ActRow, 2) & ", 'Cancelled Vouchers', '" & myAcctDesc & "', '" & flxCV.TextMatrix(ActRow, 3) & "', '" & MyReason & "', 0, 'Final', 1, " & UserID & ", #" & Now & "#);"
    'Update Auditrans
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                    "VALUES (#" & Date & "#, 'CANCEL', 'CASH VOUCHER', '" & flxCV.TextMatrix(ActRow, 2) & "', #" & flxCV.TextMatrix(ActRow, 1) & "#, '" & Left$(MyReason, 200) & "', " & UserID & ");"
    flxCV.TextMatrix(ActRow, 6) = "0.00"
    flxCV.TextMatrix(ActRow, 7) = "CANCELLED"
    MsgBox "Voucher Cancelled!", vbInformation + vbOKOnly
    
        
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 5 Then     'last index + 1
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 2 Then
            tmrDropRev.Enabled = False
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 2 'Index of Button view grid
            For i = 2 To 4
                cmdTrans(i).Top = -375
                cmdTrans(i).Left = IIf(i = 2, 0, cmdTrans(2).Width * (i - 2))
                cmdTrans(i).Visible = True
            Next i
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 1 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            cmdTrans(0).Visible = True
            cmdTrans(1).Visible = True
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub updYear_Change()
    txtYear.Text = updYear.Value
    Call cboMonth_Click
End Sub
