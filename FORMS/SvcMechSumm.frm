VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SvcMechSumm 
   Caption         =   "SERVICE CENTER"
   ClientHeight    =   8505
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFSCamt 
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Top             =   8880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSVC 
      Height          =   7095
      Left            =   120
      TabIndex        =   17
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
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   5
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
         Image           =   "SvcMechSumm.frx":0000
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8280
         TabIndex        =   7
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
         Image           =   "SvcMechSumm.frx":059A
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39322
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
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
         Image           =   "SvcMechSumm.frx":0B34
         cBack           =   -2147483633
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
         Image           =   "SvcMechSumm.frx":10CE
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
         TabIndex        =   14
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Show J.O."
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
      Image           =   "SvcMechSumm.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   11
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SvcMechSumm.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   12
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SvcMechSumm.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   13
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
      Image           =   "SvcMechSumm.frx":2736
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
      m_MDI_Height    =   9015
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic's Income Summary"
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
      Width           =   2900
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
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic's Income Summary"
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
   Begin VB.Menu mnuSPinv 
      Caption         =   "Requisition Summary"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel P.O."
      End
   End
End
Attribute VB_Name = "SvcMechSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSvc As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer

Private Sub RefreshRecord(dCode As Integer)
    Dim Qry As String, i As Integer
    
    If dCode = 1 Then
        Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
              "Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
    Else
        Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
              "Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
    End If
    If rsSvc.State = adStateOpen Then rsSvc.Close
    rsSvc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSvc.RecordCount = 0 Then
        flxSVC.Clear
        flxSVC.Rows = 2
        flxSVC.Cols = 8
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSVC.DataSource = rsSvc
    'flxSVC2.Visible = False
    Call RXResizer1_ResizeComplete
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSVC
        For i = fRow To .Rows - 1
            For j = fCol To 7
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 7:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 7, i + 1, i)
                    fCol = IIf(j = 7, 1, j + 1)
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
    Dim Qry As String, JOid As String
    
    Select Case Index
        Case 0  'Show Invoice
            If flxSVC.Rows = 2 Then ActRow = 1
            If rsSvc.RecordCount > 0 And ActRow <> 0 Then
                Qry = "SELECT SRVCJOBORDER.JOBORDERID From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxSVC.TextMatrix(ActRow, 2) & "'));"
                JOid = GetFieldValue(Qry)
                Call JobOrder.DisplayJobOrder(JOid)
                'Call POCust.ShowRecord(Val(flxSVC.TextMatrix(ActRow, 8)))
            Else
                JobOrder.Show
            End If
            Unload Me
        Case 1  'Refresh
            Call RefreshRecord(Val(cmdNext.Tag))
        Case 2  'Print
            If Val(cmdNext.Tag) = 1 Then
                Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                      "Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            Else
                Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                      "Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            End If
            
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 1, "", " to " & Format(dTo.Value, "Mmm. dd, yyyy"))
            PRNmc.ParamValue2 = Val(txtFSCamt.Text)
            PRNmc.MySource = "SRVCMECH"
            PRNmc.Show
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

Private Sub flxSVC_Click()
    ActRow = flxSVC.Row
End Sub

Private Sub flxSVC_EnterCell()
    ActRow = flxSVC.Row
End Sub

Private Sub flxSVC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If flxSVC.TextMatrix(ActRow, 6) <> "" Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuSPinv)
    End If
End Sub

Private Sub Form_Load()
    Dim Qry As String
    
    Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='FSC AMOUNT'));"
    
    txtFSCamt.Text = GetFieldValue(Qry)
    dFrom.Value = Date
    dTo.Value = Date
    cmdNext.Tag = "1"
    Call RefreshRecord(1)
    ActRow = 0
    MotoRep.sBar.Panels(6).Text = "Mechanic's Income Summary"
    MotoRep.tbMain.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsSvc.State = adStateOpen Then rsSvc.Close
End Sub

Function GetReqID(MyRN As String) As Long
    Dim rsReq As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT REQUISITION.REQID From Requisition WHERE (((REQUISITION.REQNO)='" & MyRN & "'));"
    With rsReq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetReqID = !ReqID
        Else
            GetReqID = 0
        End If
        .Close
    End With
End Function

Private Sub mnuCancel_Click()
    Dim rsSPitr As New ADODB.Recordset
    Dim MyAcc As String, MyReason As String, Qry As String
    Dim ReqID As Long, Ans As Integer, i As Integer
        
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=40));"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    'Validate  if existing invoice is cancelled
    If flxSVC.TextMatrix(ActRow, 5) <> "" Then
        Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxSVC.TextMatrix(ActRow, 5) & "') or ((SPINVOICE.REMARKS)<>'CANCELLED'));"
        If CtrlNumExist(Qry) Then
            MsgBox "Unable to cancel P.O. because an Invoice was alread issued!", vbExclamation + vbOKOnly, "Cancel Failed"
            Exit Sub
        End If
    End If
    If Val(flxSVC.TextMatrix(ActRow, 7)) = 0 Then
        Ans = MsgBox("P.O. No. " & flxSVC.TextMatrix(ActRow, 2) & " is not yet printed!" & vbCrLf & "You can still edit the P.O." & vbCrLf & vbCrLf & _
              "Are you sure you want to cancel P.O.?", vbQuestion + vbYesNo, "Cancel")
        If Ans = vbNo Then Exit Sub
    End If
    Ans = MsgBox("Do you want to Cancel P.O.?" & vbCrLf & vbCrLf & _
        "P.O. Date : " & flxSVC.TextMatrix(ActRow, 1) & vbCrLf & "P.O. No: " & flxSVC.TextMatrix(ActRow, 2) & vbCrLf & _
        "Charge to: " & flxSVC.TextMatrix(ActRow, 3), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Requisition:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Requisition failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update Requisition
    LYRcon.Execute "UPDATE CUSTPO SET CUSTPO.REMARKS = 'CANCELLED ' & [CUSTPO].[REMARKS] WHERE (((CUSTPO.CPOID)=" & flxSVC.TextMatrix(ActRow, 8) & "));"

    'Update Grid
    flxSVC.TextMatrix(ActRow, 6) = "CANCELLED"
    With flxSVC
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 150, 150)
        .Col = 2
    End With
    MsgBox "P.O. is now Cancelled!", vbInformation + vbOKOnly, "P.O. Summary"

End Sub


Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxSVC
        .FormatString = "   |<MECHANIC NAME|^J.O. NO.|^J.O. DATE|<JOB ORDER TYPE|>J.O. AMOUNT|>FSC AMOUNT|>TOTAL "
        .ColWidth(0) = 300
        .ColWidth(1) = .Width * 0.2
        .ColWidth(2) = .Width * 0.12
        .ColWidth(3) = .Width * 0.12
        .ColWidth(4) = .Width * 0.17
        .ColWidth(5) = .Width * 0.11
        .ColWidth(6) = .Width * 0.11
        .ColWidth(7) = .Width * 0.11
        If rsSvc.RecordCount > 0 Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, 3) = Format$(.TextMatrix(i, 3), "mm/dd/yyyy")
                .TextMatrix(i, 5) = Format$(.TextMatrix(i, 5), "#,##0.00")
                If InStr(.TextMatrix(i, 4), "FSC") > 0 Then .TextMatrix(i, 6) = txtFSCamt.Text
                .TextMatrix(i, 7) = Format(Val(Replace(.TextMatrix(i, 5), ",", "")) + Val(Replace(.TextMatrix(i, 6), ",", "")), "#,##0.00")
            Next i
        End If
    End With
    
    cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
    'Gillamac Information Technology Center, Inc. 169 Sanciangko St. Cebu City
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

