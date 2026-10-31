VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form arCheckSumm 
   Caption         =   "Checks Summary"
   ClientHeight    =   9060
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   10200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid flxChk 
      Height          =   6615
      Left            =   120
      TabIndex        =   18
      Tag             =   "XYF"
      Top             =   1920
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   14
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"arCheckSumm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   8520
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   8670
      Visible         =   0   'False
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtYr 
      Height          =   285
      Left            =   7800
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   7800
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6360
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   14175
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   172
         Width           =   3495
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   12120
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   10800
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
         Image           =   "arCheckSumm.frx":00D0
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   585
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   143589377
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Tag             =   "XYF"
         ToolTipText     =   "Execute"
         Top             =   585
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
         ImgAlign        =   4
         Image           =   "arCheckSumm.frx":066A
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3360
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   585
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   143589377
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   4920
         TabIndex        =   23
         Tag             =   "XYF"
         ToolTipText     =   "Execute"
         Top             =   585
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
         ImgAlign        =   4
         Image           =   "arCheckSumm.frx":0C04
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   4845
         TabIndex        =   25
         Tag             =   "XY"
         ToolTipText     =   "Execute"
         Top             =   172
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
         Image           =   "arCheckSumm.frx":119E
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "Check Date"
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
         TabIndex        =   24
         Tag             =   "xy"
         Top             =   580
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Select Status"
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
         TabIndex        =   10
         Tag             =   "xy"
         Top             =   172
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "arCheckSumm.frx":1738
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   8
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Save to Excel"
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
      Image           =   "arCheckSumm.frx":1CD2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   9
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
      Image           =   "arCheckSumm.frx":226C
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   9720
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Date   |<PO Number    |<Orders"
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&View Check"
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
      Image           =   "arCheckSumm.frx":2806
      cBack           =   -2147483633
   End
   Begin VB.Label lblTotal 
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
      Left            =   13680
      TabIndex        =   16
      Tag             =   "XY"
      Top             =   8670
      Width           =   615
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7200
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14670
      m_MDI_Height    =   9630
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK SUMMARY"
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
      Left            =   9700
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
      Caption         =   "CHECK SUMMARY"
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
      Left            =   9735
      TabIndex        =   3
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuCHK 
      Caption         =   "Check Summary"
      Visible         =   0   'False
      Begin VB.Menu mnu4Deposit 
         Caption         =   "Change Status to ""FOR DEPOSIT"""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCleared 
         Caption         =   "Change Status to ""PDC DUE"""
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Change Status to ""CANCELLED"""
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDistrib 
         Caption         =   "Show Payment Distribution"
      End
   End
End
Attribute VB_Name = "arCheckSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCRPT As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer

Private Sub cboStatus_Click()
On Error GoTo er
    Dim DateCriteria As String
    Dim strFrom As String
    Dim strTo As String
    
    Dim rsChk As New ADODB.Recordset
    Dim Qry As String, MyStat As String, Pdc4Dep As String
    Dim iPd As Integer, Ans As Integer
    Dim sKey As String, i As Integer
    
    
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
    
    If Trim(strFrom) <> "" Then strFrom = modAllocation.getMySQLDateFormat(strFrom)
    If Trim(strTo) <> "" Then strTo = modAllocation.getMySQLDateFormat(strTo)
    If strFrom = strTo Or Trim(strTo) = "" Then
        DateCriteria = " And Date(CHECKS.CHECKDATE) = Date('" & strFrom & "')"
    Else
        DateCriteria = " And Date(CHECKS.CHECKDATE) Between Date('" & strFrom & "')  and Date('" & strTo & "')"
    End If
    
    
    If cboStatus.ListIndex = -1 Then Exit Sub
    Select Case cboStatus.ListIndex
        Case 0  'For Deposit'
            sKey = "CANCELLED"
        Case 1  'Cleared
            sKey = "CURRENT-CLEARED"
        Case 2  'Cancelled
            sKey = "PDC DUE"
        Case 3  'Deposited
            sKey = "POST-DATED"
        Case 4  'Post-Dated
            sKey = "POST-DATED"
    End Select
    cmdTrans(0).Enabled = False
    cmdTrans(1).Enabled = False
    cmdTrans(2).Enabled = False
    cmdTrans(3).Enabled = False
    
    Me.Enabled = False
    
    'Display all Checks and Identify PDC checks for deposit and Deposited checks
    Qry = "SELECT CHECKS.PRDATE, CHECKS.PRNO, CHECKS.CHECKDATE, CHECKS.CHECKNO, CHECKS.CHKACCTNAME, CHECKS.CUSTOMER, CHECKS.ADDRESS, CHECKS.AMOUNT, CHECKS.BANK, CHECKS.`STATUS`, CHECKS.CHECKBAL, CHECKS.CHECKTYPE, CHECKS.CHKSPRID " & _
          "From CHECKS Where CHECKS.CHKACCTNAME IS NOT NULL AND CHECKS.`STATUS` = '" & sKey & "' " & DateCriteria & " Order By CHECKS.PRDATE DESC, CHECKS.CHECKDATE Desc;"
    
    With rsChk
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Pbar.Value = 0
            Pbar.Max = .RecordCount
            Pbar.Visible = True
            flxChk.Rows = 1
            .MoveFirst
            Pdc4Dep = ""
            Do While Not .EOF
                MyStat = IIf(IsNull(!STATUS), "", !STATUS)
                'If MyStat = "POST-DATED CHECK" And !CheckDate <= Date Then
                If InStr(MyStat, "POST-DATED") > 0 And !CHECKDATE <= Date Then
                    Pdc4Dep = IIf(Pdc4Dep <> "", Pdc4Dep & vbCrLf, "") & Format(!CHECKDATE, "mm/dd/yyyy") & vbTab & !CHECKNO & vbTab & !CHKACCTNAME & vbTab & Format(!Amount, "#,##0.00")
                    iPd = iPd + 1
                End If
                If MyStat = "FOR DEPOSIT" And !CHECKDATE < Date Then
                    !STATUS = "DEPOSITED"
                    MyStat = "DEPOSITED"
                End If
                'Display in flx  1:P.R. DATE, 2:P.R. NO. 3:CHK DATE 4:CHECK NO. 5:ACCOUNT NAME 6:CUSTOMER 7:ADDRESS  8:CHECK AMT. 9:BANK 10:STATUS 11:BALANCE 12:CHECK TYPE
                flxChk.AddItem vbTab & IIf(IsNull(!PRDATE), "", Format(!PRDATE, "mm/dd/yyyy")) & vbTab & IIf(IsNull(!PRNO), "", !PRNO) & vbTab & Format(!CHECKDATE, "mm/dd/yyyy") & vbTab & !CHECKNO & vbTab & !CHKACCTNAME & vbTab & IIf(IsNull(!CUSTOMER), "", !CUSTOMER) & vbTab & IIf(IsNull(!address), "", !address) & vbTab & Format(!Amount, "#,##0.00") & vbTab & !BANK & vbTab & MyStat & vbTab & IIf(IsNull(!CHECKBAL), "", Format(!CHECKBAL, "#,##0.00")) & vbTab & IIf(IsNull(!CHECKTYPE), "", !CHECKTYPE) & vbTab & !CHKSPRID
                .MoveNext
                Pbar.Value = Pbar.Value + 1
                DoEvents
            Loop
        Else
            flxChk.Rows = 1
            flxChk.Rows = 2
        End If
        .Close
    End With
    cmdTrans(0).Enabled = True
    cmdTrans(1).Enabled = True
    cmdTrans(2).Enabled = True
    cmdTrans(3).Enabled = True
    Pbar.Visible = False
    
    Me.Enabled = True
    If iPd > 0 Then
        MsgBox "You have " & iPd & " Post-Dated Check due!" & vbCrLf & vbCrLf & _
            "DATE" & vbTab & "CHECK NO." & vbTab & "ACCOUNT NAME" & vbTab & "AMOUNT" & vbCrLf & Pdc4Dep, vbInformation + vbOKOnly, "PDC Due"
    End If
    
    
'    With flxChk
'        .Col = 1: .Row = 1: .ColSel = 13: .RowSel = .Rows - 1: .FillStyle = flexFillRepeat: .CellBackColor = vbWhite
'        .Col = 2
'        For i = 1 To .Rows - 1
'            If InStr(.TextMatrix(i, 10), sKey) > 0 Then
'                .Col = 1: .Row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(0, 180, 0)
'                .Col = 2
'            End If
'        Next i
'    End With

Exit Sub
er:
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, "Error->cboStatus_Click()"
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxChk
        For i = fRow To .Rows - 1
            For j = fCol To 10
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 13:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 10, i + 1, i)
                    fCol = IIf(j = 10, 2, j + 1)
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
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    Dto.Visible = True
    DoEvents
    If Format(CDate(Dto.Value), "MMDDYYYY") = Format(CDate(Now()), "MMDDYYYY") Then
        Dfrom.Value = DateAdd("D", -90, Now())
    End If
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
End Sub

Private Sub cmdRefresh_Click()
    If cboStatus.ListIndex = -1 Or Trim(cboStatus.Text) = "" Then
        MsgBox "Please select a check status!", vbExclamation, "Validation!"
        cboStatus.SetFocus
        Exit Sub
    End If
    Call cboStatus_Click
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim dCov As String, Qry As String
    
    Select Case Index
        Case 0  'Display Check
            With flxChk
                If ActRow = 0 Or ActRow > .Rows - 1 Then ActRow = 1
                arPR.OwnerForm = "arCheckSumm"
                arPR.CHKSPRID = Trim(.TextMatrix(ActRow, 13))
                arPR.Show
            End With
            Unload Me
        Case 1  'Filter
            For i = 0 To 2
                 With SPfilter.cboLookIn(i)
                   'Previous Code
                   '.AddItem "PART NUMBER"
                   '.AddItem "DESCRIPTION"
                   '.AddItem "BEG. INV."
                   '.AddItem "PURCHASES"
                   '.AddItem "PURCHASE RETURN"
                   '.AddItem "SALES"
                   '.AddItem "SALES RETURN"
                   '.AddItem "ADJUSTMENTS"
                   '.AddItem "INV. END"
                   '.AddItem "BACK ORDER"
                   
                   'New Code
                   .AddItem "P.R DATE"
                   .AddItem "P.R NO"
                   .AddItem "CHECK DATE"
                   .AddItem "CHECK NUMBER"
                   .AddItem "ACCOUNT NAME"
                   .AddItem "CUSTOMER"
                   .AddItem "ADDRESS"
                   .AddItem "CHECK AMOUNT"
                   .AddItem "BANK"
                   .AddItem "STATUS"
                   .AddItem "BALANCE"
                   .AddItem "CHECK TYPE"
                   
                   
                 End With
            Next i
            'SPfilter.fldNames = "SPAREPARTS.PARTNUMBER*SPAREPARTS.DESCRIPTION*SPINVENTORY.BEGINV*SPINVENTORY.PURCHASES*SPINVENTORY.PURRET*SPINVENTORY.SALES*SPINVENTORY.SALRET*SPINVENTORY.ADJ*SPINVENTORY.ENDINV*SPINVENTORY.BACKORDER"
            SPfilter.fldNames = "CHECKS.PRDATE*CHECKS.PRNO*CHECKS.CHECKDATE*CHECKS.CHECKNO*CHECKS.CHKACCTNAME*CHECKS.CUSTOMER*CHECKS.ADDRESS*CHECKS.AMOUNT*CHECKS.BANK*CHECKS.`STATUS`*CHECKS.CHECKBAL*CHECKS.CHECKTYPE"
            SPfilter.fldNdx = "23456789" 'start w/ 0
            SPfilter.fldDates = ""
            SPfilter.fldSupplier = 10
            'SPfilter.MySource = "SPINVENTORY"
            SPfilter.MySource = "arCheckSumm"
            SPfilter.Show vbModal
            
        Case 2  'Print
            Dim XLapp As Excel.Application
            Dim wbkMVR As Excel.Workbook
            Dim wksLineA1 As Excel.Worksheet
 
            Dim sSaveAsFileName
            
            With dlgDEMO
                .CancelError = True
                .InitDir = mstrLastDir
                .Flags = cdlOFNHideReadOnly
                .filename = ""
                .Filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
                .ShowOpen
                sSaveAsFileName = .filename
            End With
            
 
            Set XLapp = CreateObject("Excel.Application")
            XLapp.Visible = True
            XLapp.Workbooks.Add
            Set wbkMVR = XLapp.ActiveWorkbook
            Set wksLineA1 = wbkMVR.Sheets(1)


            wksLineA1.Name = "Checks"
                wksLineA1.Cells(1, 1).Value = "P.R. DATE"
                wksLineA1.Cells(1, 2).Value = "P.R. NUMBER"
                wksLineA1.Cells(1, 3).Value = "CHECK DATE"
                wksLineA1.Cells(1, 4).Value = "CHECK NUMBER"
                wksLineA1.Cells(1, 5).Value = "ACCOUNT NAME"
                wksLineA1.Cells(1, 6).Value = "CUSTOMER"
                wksLineA1.Cells(1, 7).Value = "ADDRESS"
                wksLineA1.Cells(1, 8).Value = "CHECK AMOUNT"
                wksLineA1.Cells(1, 9).Value = "BANK"
                wksLineA1.Cells(1, 10).Value = "CHECK STATUS"
                wksLineA1.Cells(1, 11).Value = "BALANCE"
                With FlxRpt
                    Pbar.Value = 0
                    Pbar.Max = flxChk.Rows - 1
                    For i = 1 To flxChk.Rows - 1
                        For j = 1 To 11
                            wksLineA1.Cells(i + 1, j).Value = flxChk.TextMatrix(i, j)
                        Next j
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Next i
                End With
                XLapp.DisplayAlerts = False
 
                wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
 
                wbkMVR.Close SaveChanges:=False
                XLapp.Quit
 
                Set XLapp = Nothing
                Set wbkMVR = Nothing
                Set wksLineA1 = Nothing
                MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, "Sales DB"
        Case 3  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
    End Select
End Sub


Private Sub flxChk_DblClick()
    Call cmdTrans_Click(0) ' View Check
End Sub

Private Sub flxChk_EnterCell()
    ActRow = flxChk.Row
End Sub

Private Sub flxChk_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        With flxChk
            mnuDistrib.Enabled = False
            If Val(.TextMatrix(ActRow, 9)) > 0 Then
                If Val(.TextMatrix(ActRow, 9)) > 0 And (InStr(.TextMatrix(ActRow, 10), "PDC DUE") Or InStr(.TextMatrix(ActRow, 10), "CLEARED")) Then mnuDistrib.Enabled = True
            End If
            mnu4Deposit.Enabled = False
            mnuCleared.Enabled = False
            If InStr(.TextMatrix(ActRow, 10), "POST-DATED") > 0 Then mnuCleared.Enabled = True
            'If InStr(.TextMatrix(ActRow, 10), "DEPOSITED") > 0 Then mnuCleared.Enabled = True
        End With
        Call PopupMenu(mnuCHK)
    End If
End Sub

Private Sub Form_Load()
    Dfrom.MaxDate = Now()
    Dfrom.Value = Now()
    
    Dto.MaxDate = Now()
    Dto.Value = Now()
    
    With cboStatus
        .AddItem "CANCELLED CHECKS"
        .AddItem "CURRENT-CLEARED CHECKS"
        .AddItem "PDC DUE CHECKS"
        .AddItem "POST-DATED CHECKS"
    End With
                 
        
    Call RXResizer1_ResizeComplete
    DoEvents
    
    MotoRep.sBar.Panels(6).Text = "Summary of Checks"
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
    ActRow = 1
    DoEvents
    Timer1.Enabled = True
End Sub
                                                                        
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub mnu4Deposit_Click()
    Dim Ans As Integer, Qry As String, MyAcc As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=63));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=63;"
  
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to change Check status to FOR DEPOSIT?", vbQuestion + vbYesNo, "For Deposit")
    If Ans = vbNo Then Exit Sub
    'LYRcon.Execute "UPDATE CHECKS SET CHECKS.STATUS = 'FOR DEPOSIT', CHECKS.DATE4DPSIT = #" & Date & "# WHERE (((CHECKS.CHKSPRID)=" & Val(flxChk.TextMatrix(ActRow, 13)) & "));"
    'mySQL
    LYRcon.Execute "UPDATE CHECKS SET CHECKS.STATUS = 'FOR DEPOSIT', CHECKS.DATE4DPSIT = '" & Format(Date, "yyyy-mm-dd") & "' WHERE CHECKS.CHKSPRID=" & Val(flxChk.TextMatrix(ActRow, 13)) & ";"
  
    flxChk.TextMatrix(ActRow, 10) = "FOR DEPOSIT"
    MsgBox "Check status successfully changed!", vbInformation + vbOKOnly, "For Deposit"
    
End Sub

Private Sub mnuCancel_Click()
    Dim Ans As Integer, Qry As String, MyAcc As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=63));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=63;"

    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to change Check status to CANCELLED?", vbQuestion + vbYesNo, "CANCELLED")
    If Ans = vbNo Then Exit Sub
    'LYRcon.Execute "UPDATE CHECKS SET CHECKS.STATUS = 'CANCELLED' WHERE (((CHECKS.CHKSPRID)=" & Val(flxChk.TextMatrix(ActRow, 13)) & "));"
    'mySQL
    LYRcon.Execute "UPDATE CHECKS SET CHECKS.STATUS = 'CANCELLED' WHERE CHECKS.CHKSPRID=" & Val(flxChk.TextMatrix(ActRow, 13)) & ";"
 
    flxChk.TextMatrix(ActRow, 10) = "CANCELLED"
    MsgBox "Check status successfully changed!", vbInformation + vbOKOnly, "CANCELLED"
    
End Sub

Private Sub mnuCleared_Click()
    Dim Ans As Integer, Qry As String, MyAcc As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=63));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=63;"
 
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to change Check status to PDC DUE?", vbQuestion + vbYesNo, "Cleared")
    If Ans = vbNo Then Exit Sub
    'LYRcon.Execute "UPDATE CHECKS SET CHECKS.STATUS = 'CLEARED' WHERE (((CHECKS.CHKSPRID)=" & Val(flxChk.TextMatrix(ActRow, 13)) & "));"
    'mySQL
    LYRcon.Execute "UPDATE CHECKS SET CHECKS.STATUS = 'PDC DUE', CHECKS.UPDATEDBY = " & UserID & " WHERE CHECKS.CHKSPRID=" & Val(flxChk.TextMatrix(ActRow, 13)) & ";"
   
    flxChk.TextMatrix(ActRow, 10) = "PDC DUE"
    MsgBox "Check status successfully changed!", vbInformation + vbOKOnly, "Cleared"

End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
        
    'Display in flx  1:P.R. DATE, 2:P.R. NO. 3:CHK DATE 4:CHECK NO. 5:ACCOUNT NAME 6:CUSTOMER 7:ADDRESS  8:CHECK AMT. 9:BANK 10:STATUS 11:BALANCE 12:CHECK TYPE 13:chkprid
    With flxChk             '   1       2          3               4          5         6           7          8           9        10         11       12       13
        .ColWidth(0) = 350
        .ColWidth(1) = 0
        For i = 1 To 13
            .ColWidth(i) = (.Width - 400) * 0.09
        Next i
        For i = 5 To 7
            .ColWidth(i) = (.Width - 300) * 0.15
        Next i
        .ColWidth(9) = (.Width - 300) * 0.12
        .ColWidth(10) = (.Width - 300) * 0.15
        .ColWidth(13) = 0
    End With
    
    txtFind.Top = cmdFind.Top
    
        
    DoEvents
    Call cmdNext_Click
End Sub

Private Sub Timer1_Timer()
On Error GoTo er
    Dim rsChk As New ADODB.Recordset
    Dim Qry As String, MyStat As String, Pdc4Dep As String
    Dim iPd As Integer, Ans As Integer

    If Timer1.Enabled = False Then Exit Sub
    
    Timer1.Enabled = False
    Me.Enabled = False
    
    'Display all Checks and Identify PDC checks for deposit and Deposited checks
    Qry = "SELECT CHECKS.PRDATE, CHECKS.PRNO, CHECKS.CHECKDATE, CHECKS.CHECKNO, CHECKS.CHKACCTNAME, CHECKS.CUSTOMER, CHECKS.ADDRESS, CHECKS.AMOUNT, CHECKS.BANK, CHECKS.`STATUS`, CHECKS.CHECKBAL, CHECKS.CHECKTYPE, CHECKS.CHKSPRID " & _
          "From CHECKS Where CHECKS.CHKACCTNAME IS NOT NULL AND (CHECKS.`STATUS` = 'POST-DATED' OR CHECKS.`STATUS` = 'PDC DUE') AND CHECKS.CHECKBAL > 0 Order By CHECKS.PRDATE DESC, CHECKS.CHECKDATE Desc;"
    With rsChk
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxChk.Rows = 1
            .MoveFirst
            Pbar.Value = 0
            Pbar.Max = .RecordCount
            Pbar.Visible = True
            Pdc4Dep = ""
            Do While Not .EOF
                MyStat = IIf(IsNull(!STATUS), "", !STATUS)
                'If MyStat = "POST-DATED CHECK" And !CheckDate <= Date Then
                If InStr(MyStat, "POST-DATED") > 0 And !CHECKDATE <= Date Then
                    Pdc4Dep = IIf(Pdc4Dep <> "", Pdc4Dep & vbCrLf, "") & Format(!CHECKDATE, "mm/dd/yyyy") & vbTab & !CHECKNO & vbTab & !CHKACCTNAME & vbTab & Format(!Amount, "#,##0.00")
                    iPd = iPd + 1
                End If
                If MyStat = "FOR DEPOSIT" And !CHECKDATE < Date Then
                    !STATUS = "DEPOSITED"
                    MyStat = "DEPOSITED"
                End If
                'Display in flx  1:P.R. DATE, 2:P.R. NO. 3:CHK DATE 4:CHECK NO. 5:ACCOUNT NAME 6:CUSTOMER 7:ADDRESS  8:CHECK AMT. 9:BANK 10:STATUS 11:BALANCE 12:CHECK TYPE
                flxChk.AddItem vbTab & IIf(IsNull(!PRDATE), "", Format(!PRDATE, "mm/dd/yyyy")) & vbTab & IIf(IsNull(!PRNO), "", !PRNO) & vbTab & Format(!CHECKDATE, "mm/dd/yyyy") & vbTab & !CHECKNO & vbTab & !CHKACCTNAME & vbTab & IIf(IsNull(!CUSTOMER), "", !CUSTOMER) & vbTab & IIf(IsNull(!address), "", !address) & vbTab & Format(!Amount, "#,##0.00") & vbTab & !BANK & vbTab & MyStat & vbTab & IIf(IsNull(!CHECKBAL), "", Format(!CHECKBAL, "#,##0.00")) & vbTab & IIf(IsNull(!CHECKTYPE), "", !CHECKTYPE) & vbTab & !CHKSPRID
                .MoveNext
                Pbar.Value = Pbar.Value + 1
                DoEvents
            Loop
        End If
        .Close
    End With
    Pbar.Visible = False
    If iPd > 0 Then
        MsgBox "You have " & iPd & " Post-Dated Check due!" & vbCrLf & vbCrLf & _
            "DATE" & vbTab & "CHECK NO." & vbTab & "ACCOUNT NAME" & vbTab & "AMOUNT" & vbCrLf & Pdc4Dep, vbInformation + vbOKOnly, "PDC Due"
    End If
    
    Me.Enabled = True
Exit Sub
er:
    Me.Enabled = True
    Timer1.Enabled = False
    MsgBox Err.Description, vbCritical, "Error->Timer1_Timer()"
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Public Sub ShowFilteredRecords()
On Error GoTo er
    Dim Qry1 As String, Qry2 As String
    
    'Previous Code
    'Add Criteria in default Select
    'Qry1 = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And ((SPInventory.Month) = " & txtMo.Text & " And ((SPInventory.Year) = " & txtYr.Text & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    'Qry2 = "SELECT SPINVENTORY.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER, SPINVENTORY.MONTH, SPINVENTORY.YEAR " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And ((SPInventory.Month) = " & txtMo.Text & " And ((SPInventory.Year) = " & txtYr.Text & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    'Qry1 = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And SPInventory.Month = " & Val(txtMo.Text) & " And SPInventory.Year = " & Val(txtYr.Text) & " ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    'Qry2 = "SELECT SPINVENTORY.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPINVENTORY.BEGINV, SPINVENTORY.PURCHASES, SPINVENTORY.PURRET, SPINVENTORY.SALES, SPINVENTORY.SALRET, SPINVENTORY.ADJ, SPINVENTORY.ENDINV, SPINVENTORY.BACKORDER, SPINVENTORY.MONTH, SPINVENTORY.YEAR " & _
                  "FROM SPINVENTORY INNER JOIN SPAREPARTS ON SPINVENTORY.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And SPInventory.Month = " & Val(txtMo.Text) & " And SPInventory.Year = " & Val(txtYr.Text) & " ORDER BY SPAREPARTS.PARTNUMBER;"
    'Assign Complete Query back to txtSQL
    
    'New Code
    Qry1 = "SELECT checks.PRDATE, checks.PRNO, checks.CHECKDATE, checks.CHECKNO, checks.CHKACCTNAME, checks.CUSTOMER, checks.ADDRESS, checks.AMOUNT, checks.BANK, checks.`STATUS`, checks.CHECKBAL, checks.CHECKTYPE " & _
                "From checks " & txtSQL.Text & " ORDER BY CHECKS.PRDATE DESC"
    
    txtSQL.Text = Qry1
    
    If rsCRPT.State = adStateOpen Then rsCRPT.Close
    rsCRPT.Open txtSQL.Text, LYRcon, adOpenKeyset, adLockOptimistic
    If rsCRPT.RecordCount = 0 Then
        flxChk.Clear
        flxChk.Rows = 2
        flxChk.Cols = 12
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxChk.DataSource = rsCRPT
    ActRow = 1
    Call RXResizer1_ResizeComplete
    
Exit Sub
er:
        MsgBox Err.Description, vbCritical, "Error->ShowFilteredRecords()"
End Sub
