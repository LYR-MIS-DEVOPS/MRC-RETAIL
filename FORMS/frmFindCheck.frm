VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmFindCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Check"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15150
   Icon            =   "frmFindCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   15150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   120
      Width           =   3735
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&FIND NEXT"
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
      Image           =   "frmFindCheck.frx":57E2
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFind 
      Height          =   6135
      Left            =   120
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   600
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   15
      RowHeightMin    =   285
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin lvButton.lvButtons_H cmdSelect 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&SELECT"
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
      ImgAlign        =   3
      Image           =   "frmFindCheck.frx":5D7C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   13920
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&CANCEL"
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
      Image           =   "frmFindCheck.frx":6316
      cBack           =   -2147483633
   End
   Begin VB.Label lblSearch 
      Caption         =   "SEARCH"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Tag             =   "XYF"
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmFindCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flxFindTTLCol As Integer
Dim flxFindRowNo As Integer
Dim flxFindColNo As Integer

Dim CUSTOMERDID As Integer

Private Sub cmdCancel_Click()
On Error Resume Next
    Unload Me
    Set frmFindCheck = Nothing
End Sub

Private Sub cmdFind_Click()
    Call modTools.flxFind(flxFind, txtSearch, Me.Caption, flxFindTTLCol, flxFindRowNo, flxFindColNo)
End Sub

Sub flxFindColsChanger()
On Error Resume Next
With flxFind

    flxFindTTLCol = 15
    .Clear
    .Rows = 2
    .FixedCols = 1
    
    '               0       1           2           3           4           5              6       7                    8               9           10          11          12          13
    .FormatString = " |<CHECK NO.|<CUSTOMER NAME |^PR. NO. |<PR. DATE |<USABLE BALANCE |<BANK |<CHECK ACCOUNT NAME |<CHECK STATUS |<CHECK TYPE |<CHECK DATE |<TTL DEBIT |<TTL CREDIT |<CHECK BALANCE "
    .Col = flxFindTTLCol
    .ColWidth(0) = 200
    .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15   'CHECK NO.
    .ColWidth(2) = (.Width - .ColWidth(0)) * 0.3    'CUSTOMER NAME
    .ColWidth(3) = (.Width - .ColWidth(0)) * 0.15   'PR. NO.
    .ColWidth(4) = (.Width - .ColWidth(0)) * 0.15   'PR. DATE
    .ColWidth(5) = (.Width - .ColWidth(0)) * 0.2    'USABLE BALANCE
    .ColWidth(6) = (.Width - .ColWidth(0)) * 0.2    'BANK
    .ColWidth(7) = (.Width - .ColWidth(0)) * 0.3    'CHECK ACCOUNT NAME
    .ColWidth(8) = (.Width - .ColWidth(0)) * 0.3    'CHECK STATUS
    .ColWidth(9) = (.Width - .ColWidth(0)) * 0.3    'CHECK TYPE
    .ColWidth(10) = (.Width - .ColWidth(0)) * 0.15    'CHECK DATE
    .ColWidth(11) = (.Width - .ColWidth(0)) * 0.2    'TTL DEBIT
    .ColWidth(12) = (.Width - .ColWidth(0)) * 0.2    'TTL CREDIT
    .ColWidth(13) = (.Width - .ColWidth(0)) * 0.2   'CHECK BAL.
    .ColWidth(14) = 0  'CHKSPRID.
End With

End Sub

Sub ShowRecord()
On Error GoTo er
    Dim tmpRs As New ADODB.Recordset
    Dim i As Integer
        Call flxFindColsChanger
        
        DoEvents
        'CHKSPRID, CUSTOMERNAME, PRNO, PRDATE, BANK, CHKACCTNAME, STATUS, CHECKTYPE, CHECKNO, CHECKDATE, USABLEBAL, TTLDEBIT, TTLCREDIT, CHECKBAL
        '                 0      1               2          3        4           5                6             7          8                  9           10            11          12             13
        'FormatString = "'|<CHECK NO.|<CUSTOMER NAME |^PR. NO. |<PR. DATE |<USABLE BALANCE |<BANK |<CHECK ACCOUNT NAME |<CHECK STATUS |<CHECK TYPE |<CHECK DATE |<TTL DEBIT |<TTL CREDIT |<CHECK BALANCE "
        Call modAllocation.sp_checks_list(SysBranchID, CUSTOMERDID, tmpRs)
        If tmpRs.State = 0 Then Exit Sub
        If tmpRs.RecordCount <= 0 Then
            Exit Sub
        Else
           tmpRs.MoveFirst
           While Not tmpRs.EOF
               If Trim(flxFind.TextMatrix(1, 1)) = "" Then
                   flxFind.TextMatrix(flxFind.Row, 1) = tmpRs!CHECKNO
                   flxFind.TextMatrix(flxFind.Row, 2) = tmpRs!CustomerName
                   flxFind.TextMatrix(flxFind.Row, 3) = tmpRs!PRNO
                   flxFind.TextMatrix(flxFind.Row, 4) = tmpRs!PRDATE
                   flxFind.TextMatrix(flxFind.Row, 5) = Format(tmpRs!USABLEBAL, "#,##0.00")
                   flxFind.TextMatrix(flxFind.Row, 6) = Format(tmpRs!BANK, "#,##0.00")
                   flxFind.TextMatrix(flxFind.Row, 7) = tmpRs!CHKACCTNAME
                   flxFind.TextMatrix(flxFind.Row, 8) = tmpRs!STATUS
                   flxFind.TextMatrix(flxFind.Row, 9) = tmpRs!CHECKTYPE
                   flxFind.TextMatrix(flxFind.Row, 10) = tmpRs!CHECKDATE
                   flxFind.TextMatrix(flxFind.Row, 11) = Format(tmpRs!TTLDEBIT, "#,##0.00")
                   flxFind.TextMatrix(flxFind.Row, 12) = Format(tmpRs!TTLCREDIT, "#,##0.00")
                   flxFind.TextMatrix(flxFind.Row, 13) = Format(tmpRs!CHECKBAL, "#,##0.00")
                   flxFind.TextMatrix(flxFind.Row, 14) = tmpRs!CHKSPRID
               Else
                   flxFind.AddItem vbTab & tmpRs!CHECKNO & vbTab & tmpRs!CustomerName & vbTab & tmpRs!PRNO & vbTab & _
                   tmpRs!PRDATE & vbTab & Format(tmpRs!USABLEBAL, "#,##0.00") & vbTab & _
                   Format(tmpRs!BANK, "#,##0.00") & vbTab & tmpRs!CHKACCTNAME & vbTab & _
                   tmpRs!STATUS & vbTab & tmpRs!CHECKTYPE & vbTab & tmpRs!CHECKDATE & vbTab & _
                   Format(tmpRs!TTLCREDIT, "#,##0.00") & vbTab & Format(tmpRs!TTLDEBIT, "#,##0.00") & vbTab & _
                   Format(tmpRs!CHECKBAL, "#,##0.00") & vbTab & tmpRs!CHKSPRID
                  
               End If
               tmpRs.MoveNext
            Wend
        End If

Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->ShowRecord()"
End Sub

Private Sub cmdSelect_Click()
On Error GoTo er
    Dim CHECKNO As String
    Dim CustomerName As String
    Dim PRNO As String
    Dim PRDATE As String
    Dim USABLEBAL As String
    Dim BANK As String
    Dim STATUS As String
    Dim CHECKTYPE As String
    Dim CHECKDATE As String
    Dim CHECKBAL As String
    Dim CHKSPRID As Integer
    
    If cmdSelect.Visible = False Then Exit Sub
    
    If Trim(flxFind.TextMatrix(flxFind.Row, 1)) = "" Then
        MsgBox "PLEASE SELECT A RECORD!", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    CHECKNO = flxFind.TextMatrix(flxFind.Row, 1)
    CustomerName = flxFind.TextMatrix(flxFind.Row, 2)
    PRNO = flxFind.TextMatrix(flxFind.Row, 3)
    PRDATE = flxFind.TextMatrix(flxFind.Row, 4)
    USABLEBAL = flxFind.TextMatrix(flxFind.Row, 5)
    BANK = flxFind.TextMatrix(flxFind.Row, 6)
    STATUS = flxFind.TextMatrix(flxFind.Row, 8)
    CHECKTYPE = flxFind.TextMatrix(flxFind.Row, 9)
    CHECKDATE = flxFind.TextMatrix(flxFind.Row, 10)
    CHECKBAL = flxFind.TextMatrix(flxFind.Row, 13)
    CHKSPRID = flxFind.TextMatrix(flxFind.Row, 14)
    
    '               0    1               2           3         4           5             6              7             8                   9          10          11          12             13
    'FormatString = "'|<CHECK NO.|<CUSTOMER NAME |^PR. NO. |<PR. DATE |<USABLE BALANCE |<BANK |<CHECK ACCOUNT NAME |<CHECK STATUS |<CHECK TYPE |<CHECK DATE |<TTL DEBIT |<TTL CREDIT |<CHECK BALANCE
    Call frmAllocation.ShowRecord(CHECKNO, CustomerName, PRNO, PRDATE, USABLEBAL, BANK, STATUS, CHECKTYPE, CHECKDATE, CHECKBAL, CHKSPRID)
    Call frmAllocation.LockedChckAllocDetail
    Call frmAllocation.ComputeNewCheckBalance
        
    DoEvents
    Call frmAllocation.ComputeNewCheckBalance
    Unload Me
    Set frmFindCheck = Nothing
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->cmdSelect_Click()"
    DoEvents
    Unload Me
    Set frmFindCheck = Nothing
End Sub

Private Sub flxFind_DblClick()
On Error GoTo er

    Call cmdSelect_Click

Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->flxFind_DblClick()"
End Sub

Private Sub Form_Load()
On Error GoTo er
    CUSTOMERDID = -1
    
    Call ShowRecord
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->Form_Load()"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
On Error GoTo er
    
    If KeyAscii = 13 Then
        Call cmdFind_Click
    End If
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->Form_Load()"
End Sub
