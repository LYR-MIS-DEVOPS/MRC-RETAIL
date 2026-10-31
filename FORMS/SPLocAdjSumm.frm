VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPLocAdjSumm 
   Caption         =   "Summary of Adjusted Location"
   ClientHeight    =   8550
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid FlxPrev 
      Height          =   1095
      Left            =   1680
      TabIndex        =   17
      Top             =   8880
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      FormatString    =   "SPID|<PREVIOUS LOCATION "
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxJob 
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
         Image           =   "SPLocAdjSumm.frx":0000
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
         Format          =   45613057
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   12
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
         Image           =   "SPLocAdjSumm.frx":059A
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   14
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
         Image           =   "SPLocAdjSumm.frx":0B34
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   15
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45613057
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   3240
         TabIndex        =   16
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
         Image           =   "SPLocAdjSumm.frx":10CE
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
      Caption         =   "&Adj. Location"
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
      Image           =   "SPLocAdjSumm.frx":1668
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
      Image           =   "SPLocAdjSumm.frx":1C02
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
      Image           =   "SPLocAdjSumm.frx":219C
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
      m_MDI_Height    =   9060
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SPAREPARTS LOCATION ADJUSTMENT SUMMARY"
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
      Left            =   6240
      TabIndex        =   2
      Top             =   450
      Width           =   5505
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
      Caption         =   "SPAREPARTS LOCATION ADJUSTMENT SUMMARY"
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
      Left            =   6255
      TabIndex        =   3
      Top             =   480
      Width           =   5505
   End
   Begin VB.Menu mnuAdj 
      Caption         =   "ADJ summary"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Adjustment"
      End
   End
End
Attribute VB_Name = "SPLocAdjSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsJob As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Dim Qry As String

Private Sub RefreshRecord(dCode As Integer)
    Dim i As Integer
    
    If dCode = 1 Then
        Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
              "From TRANSFERLOC Where (((TRANSFERLOC.DATETRANS) = #" & dFrom.Value & "#)) ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
    Else
        Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
              "From TRANSFERLOC Where (((TRANSFERLOC.DATETRANS) >= #" & dFrom.Value & "# And (TRANSFERLOC.DATETRANS) <= #" & dTo.Value & "#)) ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
    End If
    If rsJob.State = adStateOpen Then rsJob.Close
    rsJob.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsJob.RecordCount = 0 Then
        flxJob.Clear
        flxJob.Rows = 2
        flxJob.Cols = 8
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxJob.DataSource = rsJob
    ActRow = 1
    flxJob.Visible = True
    'flxjob2.Visible = False
    Call RXResizer1_ResizeComplete
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxJob
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
        Case 1  'ADJUST LOCATION
            With flxJob
                If ActRow = 0 And .Rows >= 2 Then ActRow = 1
                If flxJob.TextMatrix(ActRow, 8) = "" Then
                    MsgBox "No Location Adjustment Available!", vbInformation + vbOKOnly, "Location Adjustment"
                    Exit Sub
                End If
                Call SPLocAdj.DisplayLocAdj(flxJob.TextMatrix(ActRow, 8))
            End With
            Unload Me
        Case 2  'Print
            'dCov = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 2, " to " & Format(dTo.Value, "Mmm. dd, yyyy"), "")
            'If Val(cmdNext.Tag) = 1 Then
            '    Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
            '          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            'Else
            '    Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
            '          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            'End If
            'PRNmc.QryPrn = Qry
            'PRNmc.ParamValue1 = dCov
            'PRNmc.ParamValue2 = "Prepared By: " & GetUserName
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

Private Sub flxjob_EnterCell()
    ActRow = flxJob.Row
End Sub

Private Sub flxJob_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Val(flxJob.TextMatrix(ActRow, 7)) = 0 Or flxJob.TextMatrix(ActRow, 3) = "CANCELLED" Then Exit Sub
    If DateValue(flxJob.TextMatrix(ActRow, 1)) <> Date Then Exit Sub
    If Button = vbRightButton Then
       Call PopupMenu(mnuADJ)
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
    
    MotoRep.sBar.Panels(6).Text = "Summary of Adjusted Location"
    MotoRep.tbMain.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsJob.State = adStateOpen Then rsJob.Close
End Sub

Private Sub mnuCancel_Click()
    Dim Ans As Integer, MyReason As String, Qry As String
    Dim rsPRV As New ADODB.Recordset
    
    
    Ans = MsgBox("Do you want to Cancel Adjustment with Control No. " & flxJob.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Adjustment:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Job Order failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update SP Location Adjustment
    LYRcon.Execute "UPDATE TRANSFERLOC SET TRANSFERLOC.REFERENCE = 'CANCELLED' WHERE (((TRANSFERLOC.TRANSLOCID)=" & flxJob.TextMatrix(ActRow, 8) & "));"
    Qry = "SELECT TRANSLOCDTLS.SPAREPARTID, TRANSLOCDTLS.PREVLOC From TRANSLOCDTLS WHERE (((TRANSLOCDTLS.TRANSLOCID)=" & flxJob.TextMatrix(ActRow, 8) & "));"
    With rsPRV
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            FlxPrev.Rows = 1
            .MoveFirst
            Do While Not .EOF
                FlxPrev.AddItem !SPAREPARTID & vbTab & !PrevLoc
                .MoveNext
            Loop
        End If
        .Close
    End With
    With FlxPrev
        For i = 1 To .Rows - 1
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & .TextMatrix(i, 1) & "' WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 0) & "));"
        Next i
    End With
    'Update Audit Trans
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'SP LOCATION ADJ.', '" & flxJob.TextMatrix(ActRow, 2) & "', #" & flxJob.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
   'Update Grid
    flxJob.TextMatrix(ActRow, 3) = "CANCELLED"
    With flxJob     'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
    End With
    MsgBox "Spare Parts Location Adjustment Cancelled!", vbInformation + vbOKOnly

End Sub

Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxJob
        .FormatString = "   |^ DATE|^CONTROL NO.|<REFERENCE|<REASON|<PREPARED BY|<APPROVED BY|^PRINT COPY"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.2
        .ColWidth(4) = (.Width - 300) * 0.4
        .ColWidth(5) = (.Width - 300) * 0.15
        .ColWidth(6) = (.Width - 300) * 0.15
        .ColWidth(7) = (.Width - 300) * 0.1
        .ColWidth(8) = 0
        If .Rows >= 2 And .TextMatrix(1, 1) <> "" Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, 2) = Format$(.TextMatrix(i, 2), "0000")
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "MM/DD/YYYY")
                If Val(.TextMatrix(i, 7)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = .Cols - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 3) = "CANCELLED" Then
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .Col = 2
                End If
            Next i
        End If
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

