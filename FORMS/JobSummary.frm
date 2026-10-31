VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form JobSummary 
   Caption         =   "Job Order Summary"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxJO 
      Height          =   6975
      Left            =   0
      TabIndex        =   12
      Tag             =   "XY"
      Top             =   1200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"JobSummary.frx":0000
   End
   Begin CRVIEWERLibCtl.CRViewer vwrJob 
      Height          =   7695
      Left            =   0
      TabIndex        =   7
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
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   400
      Width           =   11295
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52035585
         CurrentDate     =   39055
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date:"
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
         TabIndex        =   5
         Top             =   285
         Width           =   2505
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date:"
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
         TabIndex        =   6
         Top             =   315
         Width           =   2505
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
      Caption         =   "&Show Job Order"
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
      Image           =   "JobSummary.frx":00E0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Filter Records"
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
      Image           =   "JobSummary.frx":067A
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
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
      Image           =   "JobSummary.frx":0C14
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   3
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
      Image           =   "JobSummary.frx":11AE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   8
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
      Image           =   "JobSummary.frx":1748
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   9
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
      Image           =   "JobSummary.frx":1CE2
      cBack           =   -2147483633
   End
   Begin VB.Label lblP 
      Caption         =   "x"
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
      TabIndex        =   10
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11520
      m_MDI_Height    =   8625
   End
   Begin VB.Menu mnuJob 
      Caption         =   "Job Order"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Job Order"
      End
   End
End
Attribute VB_Name = "JobSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Public ViewGrid As Boolean
Dim BtnIndx As Integer, Qry As String
Dim ActRow As Integer

Private Sub cmdTrans_Click(Index As Integer)
    Dim MyQry As String
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
            Call dtpDate_Change
            Call FlxColumnResize
        Case 1  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
        Case 2  'Show Job Order
            With flxJO
                If ActRow = 0 Or ActRow > .Rows - 1 Then
                    MsgBox "Please select a record to view!", vbInformation + vbOKOnly, "Confirm"
                    Exit Sub
                End If
                Call JobOrder.DisplayJobOrder(flxJO.TextMatrix(ActRow, 12))
            End With
            Unload Me
        Case 4  'Print Preview
            BtnIndx = 5
            tmrFloatRev.Enabled = True
            Call PrintPreView(True)
            Qry = "SELECT SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOBORDERNO, SRVCUSTOMER.CUSTOMER, SRVCJOBORDER.JOTYPE, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.DOWNPAYMENT, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCUSTOMER INNER JOIN SRVCJOBORDER ON SRVCUSTOMER.ACCOUNTNO = SRVCJOBORDER.ACCOUNTNO Where (((SRVCJOBORDER.DATERECEIVED) = #" & dtpDate.Value & "#)) ORDER BY SRVCJOBORDER.JOBORDERNO DESC;"
           ' Call DisplayReport(Qry)
        Case 5  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
    End Select
End Sub


Private Sub dtpDate_Change()
    Dim MyQry As String
    
    MyQry = "SELECT SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOBORDERNO, SRVCUSTOMER.CUSTOMER, SRVCJOBORDER.JOTYPE, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.DOWNPAYMENT, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.JOBORDERID " & _
            "FROM SRVCUSTOMER INNER JOIN SRVCJOBORDER ON SRVCUSTOMER.ACCOUNTNO = SRVCJOBORDER.ACCOUNTNO Where (((SRVCJOBORDER.DATERECEIVED) = #" & dtpDate.Value & "#)) ORDER BY SRVCJOBORDER.JOBORDERNO DESC;"
    
    If rsRPT.State = adStateOpen Then rsRPT.Close
    With rsRPT
        .Open MyQry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxJO.Rows = 1
            Do While Not .EOF
                flxJO.AddItem vbTab & Format(!DATERECEIVED, "mm/dd/yyyy") & vbTab & !JobOrderNo & vbTab & !Customer & vbTab & !JOType & vbTab & !Model & vbTab & !Engine & vbTab & Format$(!TotalCost, "###,##0.00") & vbTab & Format$(!DownPayment, "###,##0.00") & vbTab & Format$(!Balance, "###,##0.00") & vbTab & IIf(IsNull(!ORNO), "", !ORNO) & vbTab & !PrintCopy & vbTab & !JoborderID
                .MoveNext
            Loop
            ActRow = 1
        Else
            flxJO.Rows = 1
            flxJO.Rows = 2
        End If
        .Close
    End With

    Call FlxColumnResize

End Sub


Private Sub flxjo_Click()
    If flxJO.Col = 0 And flxJO.Row > 0 Then
        ActRow = flxJO.Row
        Call flxjo_EnterCell
    End If
End Sub

Private Sub flxjo_EnterCell()
    ActRow = flxJO.Row
    With flxJO
        .Row = ActRow: .Col = 0
        .RowSel = ActRow: .ColSel = .Cols - 1
    End With
End Sub

Private Sub flxjo_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim Qry As String
    If Button = vbRightButton And flxJO.TextMatrix(ActRow, 4) <> "Cancelled" And Val(flxJO.TextMatrix(ActRow, 11)) > 0 Then
        Call PopupMenu(mnuJob)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    dtpDate.Value = Date
    BtnIndx = 1
    tmrFloat.Enabled = True
    Call dtpDate_Change
    Call PrintPreView(False)
    Call FlxColumnResize
End Sub


'Private Sub DisplayReport(MyQry As String)
'    Dim crDisb As New Disbursement
'    Dim Disburser As String
'    Dim i As Integer
'
'    Disburser = GetSettings("DISBURSER")
'    If rsRPT.State = adStateOpen Then rsRPT.Close
'    rsRPT.Open MyQry, LYRcon, adOpenKeyset, adLockOptimistic
'    crDisb.DiscardSavedData
'    crDisb.Database.SetDataSource rsRPT
'    crDisb.EnableParameterPrompting = False
'    vwrJob.ReportSource = crDisb
'    vwrJob.ViewReport
'End Sub

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
        Frame1.Visible = False
        flxJO.Visible = False
        vwrJob.Visible = True
    Else
        vwrJob.Visible = False
        Frame1.Visible = True
        flxJO.Visible = True
    End If
End Sub

Private Sub FlxColumnResize()
    Dim i As Integer
    
    With flxJO
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 2500
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .ColWidth(6) = 2000
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .ColWidth(12) = 1200
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 11)) = 0 Then
                '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                .Col = 1: .Row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
            End If
        
        Next i
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
End Sub

Private Sub mnuCancel_Click()
    Dim Ans As Integer, myReason As String
    Dim myAcctDesc As String * 50
    Dim Qry As String
    Dim i As Integer
    
    'Validate job order if payment has been made
    If flxJO.TextMatrix(ActRow, 11) <> "" Then
        MsgBox "Cancellation not allowed because payment was already made!", vbExclamation + vbOKOnly, "O.R. Exist"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Cancel Job Order?" & vbCrLf & vbCrLf & _
        "Job Order Date : " & flxJO.TextMatrix(ActRow, 1) & vbCrLf & _
        "Job Order No : " & flxJO.TextMatrix(ActRow, 2) & vbCrLf & _
        "Customer Name: " & flxJO.TextMatrix(ActRow, 3) & vbCrLf & _
        "Job Order Type: " & flxJO.TextMatrix(ActRow, 4), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    myReason = InputBox("Please indicate reason for Cancellation of Job Order:", "Reason for Cancel")
    If myReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Job Order failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update Job Order
    LYRcon.Execute "UPDATE MCJOBORDER SET MCJOBORDER.JOTYPE = 'Cancelled', MCJOBORDER.CUSTOMER = Null, MCJOBORDER.ADDRESS = Null, MCJOBORDER.CONTACTNO = Null, MCJOBORDER.ACCOUNTNO = Null, MCJOBORDER.MCMODEL = Null, MCJOBORDER.CHASSIS = Null, MCJOBORDER.PLATENO = Null, MCJOBORDER.TOTALCOST = 0, MCJOBORDER.DOWNPAYMENT = 0, MCJOBORDER.BALANCE = 0, MCJOBORDER.KMREADING = Null, MCJOBORDER.REMARKS = [MCJOBORDER].[REMARKS] & ' Reason for Cancel: " & myReason & "' " & _
                   "WHERE (((MCJOBORDER.JOBORDERID)=" & flxJO.TextMatrix(ActRow, 13) & "));"
    'Update Service Income if Exist
    Qry = "SELECT SERVICEINCOME.SRVCDATE, SERVICEINCOME.PRINTCOPY From SERVICEINCOME " & _
          "WHERE (((SERVICEINCOME.SRVCDATE)=#" & dtpDate.Value & "#));"
    If CtrlNumExist(Qry) Then   'Update Service income if record exist
        LYRcon.Execute "UPDATE SERVICEINCOME SET SERVICEINCOME.CANCELJO = IIf(IsNull([SERVICEINCOME].[CANCELJO]),'" & flxJO.TextMatrix(ActRow, 2) & "',[SERVICEINCOME].[CANCELJO] & ', " & flxJO.TextMatrix(ActRow, 2) & "'), SERVICEINCOME.TTLCANCELJO = [SERVICEINCOME].[TTLCANCELJO]+1 " & _
                       "WHERE (((SERVICEINCOME.SRVCDATE)=#" & dtpDate.Value & "#));"
    End If
    'Update Grid
    For i = 3 To 10
        flxJO.TextMatrix(ActRow, i) = ""
    Next i
    flxJO.TextMatrix(ActRow, 4) = "Cancelled"
    MsgBox "Job Order Cancelled!", vbInformation + vbOKOnly

End Sub


Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then     'last index + 1
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
            For i = 2 To 5
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

