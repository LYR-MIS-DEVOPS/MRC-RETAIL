VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAllocationSummary 
   Caption         =   "ALLOCATION SUMMARY"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   13800
   Tag             =   "XY"
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   120
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   75
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
      Caption         =   "E&xit Summary"
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
      Image           =   "frmAllocationSummary.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdShow 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "Sho&w Form"
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
      Image           =   "frmAllocationSummary.frx":0A12
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSummary 
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   1350
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   11245
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
      TabIndex        =   5
      Tag             =   "XY"
      Top             =   720
      Width           =   13575
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   11520
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   10200
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
         Image           =   "frmAllocationSummary.frx":0FAC
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   138936321
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   9
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
         Image           =   "frmAllocationSummary.frx":1546
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   10
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
         Image           =   "frmAllocationSummary.frx":1AE0
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   139001857
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
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
         Image           =   "frmAllocationSummary.frx":207A
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
         TabIndex        =   13
         Top             =   210
         Width           =   1335
      End
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
      TabIndex        =   3
      Top             =   520
      Width           =   2610
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ALLOCATION SUMMARY"
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
      Index           =   1
      Left            =   9120
      TabIndex        =   1
      Top             =   520
      Width           =   4500
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14040
      m_MDI_Height    =   8385
   End
   Begin VB.Menu mnuDummy 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Allocation"
      End
   End
End
Attribute VB_Name = "frmAllocationSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fRow  As Integer
Dim fCol  As Integer
Dim MaxCol As Integer
Dim flxActiveRow As Integer
Dim ActRow As Integer

Sub flxSummaryColsChanger()
    On Error Resume Next
    With flxSummary
         MaxCol = 11
        .Clear
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .FormatString = " |ALLOCATION NO |<DATE |<CUSTOMER |>ALLOCATED AMOUNT |>TAX w/HELD |>NET PAYABLE |<STATUS |<PAY TYPE |<RESPONSIBLE |<LASTUPDATE"
        'ALLOCATIONID, ALLOCATIONDATE, CUSTOMERNAME, ALLOCATEDAMOUNT, TAXWHELDAMOUNT, NETPAYABLE, STATUS, PAYTYPE, RESPONSIBLE, LASTUPDATE
        .Cols = MaxCol
    End With
    DoEvents
    Call flxSummarySetColsWidth
End Sub

Sub flxSummarySetColsWidth()
On Error Resume Next
    With flxSummary
        '.FormatString = " |ALLOCATION NO |<DATE |<CUSTOMER |>ALLOCATED AMOUNT |<TAX w/HELD |<NET PAYABLE |<STATUS |<PAY TYPE |<RESPONSIBLE |<LASTUPDATE"
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.07 'ALLOCATION NO
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.1 'DATE
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.2 'CUSTOMER
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.09 'ALLOCATED AMOUNT
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.09 'TAX w/HELD
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.09 'NET PAYABLE
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.1 'STATUS
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.1 'PAY TYPE
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.15 'RESPONSIBLE
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.12  'LASTUPDATE
    End With
End Sub


Private Sub cmdExit_Click()
    Call Unload(frmAllocationSummary)
    Set frmAllocationSummary = Nothing
    'MotoRep.imgBG.Visible = True
    MotoRep.toolBarButton.Visible = True
End Sub

Private Sub cmdFind_Click()
On Error GoTo er
    Dim i As Integer, j As Integer
    If txtFind.Text = "" Then Exit Sub
    With flxSummary
        For i = fRow To .Rows - 1
            For j = fCol To MaxCol - 1
                If InStr(UCase(.TextMatrix(i, j)), UCase(txtFind.Text)) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = MaxCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = MaxCol - 1, i + 1, i)
                    fCol = IIf(j = MaxCol - 1, 1, j + 1)
                    flxActiveRow = i
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cmdFind_Click()"
End Sub

Private Sub cmdNext_Click()
    'cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    Dto.Visible = True
    cmdRefresh.Visible = True

    DoEvents
    If Format(CDate(Dto.Value), "MMDDYYYY") = Format(CDate(Now()), "MMDDYYYY") Then
        Dfrom.Value = DateAdd("D", -15, Now())
    End If
End Sub

Private Sub cmdPrev_Click()
    'cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    cmdRefresh.Visible = False
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim RowIndex As Integer
    
    Dim strFrom As String
    Dim strTo As String
   

    Call flxSummaryColsChanger
    
    If Dto.Visible Then
        If CDate(Dfrom.Value) > CDate(Dto.Value) Then
            MsgBox "Invalid date range!", vbExclamation, Me.Caption
            Dfrom.SetFocus
            Exit Sub
        End If
        strFrom = Dfrom.Value
        strTo = Dto.Value
    Else
        strFrom = Dfrom.Value
        strTo = ""
    End If
    DoEvents
    
    Call modAllocation.sp_allocation_summary(strFrom, strTo, tmpRs)
    
    '.FormatString = " |ALLOCATIONID |<DATE |<CUSTOMER |>ALLOCATED AMOUNT |>TAX w/HELD |>NET PAYABLE |<STATUS |<PAY TYPE |<RESPONSIBLE |<LASTUPDATE"
    'ALLOCATIONID, ALLOCATIONDATE, CUSTOMERNAME, ALLOCATEDAMOUNT, TAXWHELDAMOUNT, NETPAYABLE, STATUS, PAYTYPE, RESPONSIBLE, LASTUPDATE
    With tmpRs
        If .State = 1 Then
            If .RecordCount > 0 Then
                 .MoveFirst
                While Not .EOF
                     RowIndex = RowIndex + 1
                     DoEvents
                    If RowIndex = 1 Then
                         flxSummary.TextMatrix(1, 1) = !AllocationID
                         flxSummary.TextMatrix(1, 2) = !ALLOCATIONDATE
                         flxSummary.TextMatrix(1, 3) = !CustomerName
                         flxSummary.TextMatrix(1, 4) = IIf(!AllocatedAmount = "0.00", "", Format(!AllocatedAmount, "#,##0.00")) & "  "
                         flxSummary.TextMatrix(1, 5) = IIf(!TAXWHELDAMOUNT = "0.00", "", Format(!TAXWHELDAMOUNT, "#,##0.00")) & "  "
                         flxSummary.TextMatrix(1, 6) = IIf(!NetPayable = "0.00", "", Format(!NetPayable, "#,##0.00")) & "  "
                         flxSummary.TextMatrix(1, 7) = !STATUS
                         flxSummary.TextMatrix(1, 8) = !PayType
                         flxSummary.TextMatrix(1, 9) = !RESPONSIBLE
                         flxSummary.TextMatrix(1, 10) = !LASTUPDATE
                    Else
                         flxSummary.AddItem _
                             vbTab & !AllocationID & _
                             vbTab & !ALLOCATIONDATE & _
                             vbTab & !CustomerName & _
                             vbTab & IIf(!AllocatedAmount = "0.00", "", Format(!AllocatedAmount, "#,##0.00")) & "  " & _
                             vbTab & IIf(!TAXWHELDAMOUNT = "0.00", "", Format(!TAXWHELDAMOUNT, "#,##0.00")) & "  " & _
                             vbTab & IIf(!NetPayable = "0.00", "", Format(!NetPayable, "#,##0.00")) & "  " & _
                             vbTab & !STATUS & _
                             vbTab & !PayType & _
                             vbTab & !RESPONSIBLE & _
                             vbTab & !LASTUPDATE
                    End If
                    .MoveNext
                Wend
            End If
        End If
    End With
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cmdRefresh_Click()"
End Sub
 
Private Sub cmdShow_Click()
    frmAllocation.OwnerForm = "frmAllocationSummary"
    
    If flxSummary.Row >= 1 Then
        frmAllocation.AllocationID = flxSummary.TextMatrix(flxSummary.Row, 1)
    End If
    frmAllocation.Show vbModal
End Sub

Private Sub flxSummary_DblClick()
    Call cmdShow_Click
End Sub

Private Sub flxSummary_EnterCell()
    ActRow = flxSummary.Row
End Sub

Private Sub flxSummary_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuDummy)
    End If
End Sub

Private Sub Form_Load()
On Error GoTo er
    Dim tmpMonthNo As Integer, tmpYear As Integer
    Dim tmpDate  As Date
    tmpDate = CDate(ServerTime)
    tmpMonthNo = Month(tmpDate)
    tmpYear = Year(tmpDate)
    
    Dto.MaxDate = CDate(tmpMonthNo & "-" & modTools.GetLastDayOfTheMonth(tmpMonthNo, tmpYear) & "-" & tmpYear)
    Dfrom.MaxDate = Dto.MaxDate
    Dto.Value = Dto.MaxDate
    Dfrom.Value = CDate(tmpMonthNo & "-1-" & tmpYear)
    
    Call cmdNext_Click
    
    tmrLoad.Enabled = True
Exit Sub
er:
End Sub

Private Sub mnuCancel_Click()
Dim chkExists As Long
Dim Qry As String

    If flxSummary.TextMatrix(ActRow, 1) <> "" Then
        chkExists = Val(GetFieldValue("Select PAYMENTID From payments Where ALLOCATIONID = " & Val(flxSummary.TextMatrix(ActRow, 1)) & " "))
        If Val(chkExists) = 0 Then
            Qry = "call sp_mrcretails.sp_allocation_cancel(" & Val(SysBranchID) & ",'" & sysBranchCode & "','" & flxSummary.TextMatrix(ActRow, 1) & "','" & "CANCELLED" & "'," & Val(UserID) & ")"
            With rsAnyUse
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    If !level = "Error" Then
                        MsgBox !message, vbInformation + vbOKOnly, "CANCELLATION"
                    Else
                        MsgBox !message, vbInformation + vbOKOnly, "CANCELLATION"
                    End If
                End If
                .Close
            End With
        Else
            MsgBox "Cannot canel due to C.R / O.R is already issued!."
            Exit Sub
        End If
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
On Error Resume Next
    lblHeading(1).Left = (Me.Width - lblHeading(1).Width) - 300
    
    Call flxSummaryColsChanger
End Sub

Private Sub tmrLoad_Timer()
On Error GoTo er
    If tmrLoad.Enabled = False Then Exit Sub
    tmrLoad.Enabled = False
    Call cmdRefresh_Click
Exit Sub
er:
    tmrLoad.Enabled = False
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdFind_Click
    End If
End Sub
