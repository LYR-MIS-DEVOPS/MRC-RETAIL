VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form DPSummary 
   Caption         =   "Down Payment Summary"
   ClientHeight    =   8475
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDP 
      Height          =   7095
      Left            =   120
      TabIndex        =   10
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
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   180
         Width           =   2655
      End
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
         Image           =   "DPSummary.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Select Status:"
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
         TabIndex        =   9
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   7
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
      Image           =   "DPSummary.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   8
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
      Image           =   "DPSummary.frx":0B34
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12105
      m_MDI_Height    =   9045
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DOWN PAYMENT SUMMARY"
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
      Caption         =   "DOWN PAYMENT SUMMARY"
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
   Begin VB.Menu mnuDP 
      Caption         =   "Down Payment summary"
      Visible         =   0   'False
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show Allocation of DP"
      End
      Begin VB.Menu mnuAllocate 
         Caption         =   "Allocate Down Payment"
      End
   End
End
Attribute VB_Name = "DPSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDP As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Dim Qry As String

Private Sub RefreshRecord(MyStat As String)
    
    If MyStat = "ALL DP" Then
        Qry = "SELECT DPLEDGER.ORDATE, DPLEDGER.ORNO, CUSTOMER.CUSTOMERNAME, DPLEDGER.ORAMOUNT, DPLEDGER.STATUS, DPLEDGER.BALANCE, DPLEDGER.INVREF, DPLEDGER.JOREF, DPLEDGER.REMARKS, CUSTOMER.CUSTOMERID, DPLEDGER.AMTALLOC " & _
              "FROM DPLEDGER INNER JOIN CUSTOMER ON DPLEDGER.CUSTOMERID = CUSTOMER.CUSTOMERID ORDER BY DPLEDGER.ORDATE DESC;"
    Else
        Qry = "SELECT DPLEDGER.ORDATE, DPLEDGER.ORNO, CUSTOMER.CUSTOMERNAME, DPLEDGER.ORAMOUNT, DPLEDGER.STATUS, DPLEDGER.BALANCE, DPLEDGER.INVREF, DPLEDGER.JOREF, DPLEDGER.REMARKS, CUSTOMER.CUSTOMERID, DPLEDGER.AMTALLOC " & _
              "FROM DPLEDGER INNER JOIN CUSTOMER ON DPLEDGER.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((DPLEDGER.STATUS)='" & MyStat & "')) ORDER BY DPLEDGER.ORDATE DESC;"
    End If
    If rsDP.State = adStateOpen Then rsDP.Close
    rsDP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsDP.RecordCount = 0 Then
        flxDP.Clear
        flxDP.Rows = 2
        flxDP.Cols = 9
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxDP.DataSource = rsDP
    ActRow = 1
    flxDP.Visible = True
    Call RXResizer1_ResizeComplete
End Sub

Private Sub cboStatus_Click()
    If cboStatus.ListIndex = -1 Then Exit Sub
    Call RefreshRecord(cboStatus.Text)
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxDP
        For i = fRow To .Rows - 1
            For j = fCol To 9
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 9:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 9, i + 1, i)
                    fCol = IIf(j = 9, 1, j + 1)
                    ActRow = i
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

Private Sub cmdTrans_Click(Index As Integer)
    Dim dCov As String, Qry As String
    
    Select Case Index
        Case 2  'Print
            MsgBox "Not Yet Available", vbInformation, "Print"
        Case 3  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
    End Select
End Sub

Private Sub flxDP_EnterCell()
    ActRow = flxDP.Row
End Sub

Private Sub flxDP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InStr(UCase$(flxDP.TextMatrix(ActRow, 5)), "CANCELLED") > 0 Then Exit Sub
    If Button = vbRightButton Then
        If flxDP.TextMatrix(ActRow, 7) = "" And flxDP.TextMatrix(ActRow, 8) = "" Then
            mnuShowAll.Enabled = False
        Else
            mnuShowAll.Enabled = True
        End If
        If flxDP.TextMatrix(ActRow, 5) = "OPEN" Then
            mnuAllocate.Enabled = True
        Else
            mnuAllocate.Enabled = False
        End If
        
        Call PopupMenu(mnuDP)
    End If
End Sub

Private Sub Form_Load()
        
    cboStatus.AddItem "ALL DP"
    cboStatus.AddItem "OPEN"
    cboStatus.AddItem "CLOSED"
    'Call RefreshRecord("ALL DP")
    cboStatus.ListIndex = 0
    ActRow = 1
    MotoRep.sBar.Panels(6).Text = "Down Payment Summary"
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsDP.State = adStateOpen Then rsDP.Close
End Sub

Private Sub mnuAllocate_Click()
    Dim Qry As String, MyAcc As String
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=68));"
    'mySQL
'    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=68;"
'
'    MyAcc = GetFieldValue(Qry)
'    If MyAcc <> "FULL ACCESS" Then
'        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
'        Exit Sub
'    End If

    DPallocate.RowSrc = ActRow
    DPallocate.MySource = "ALLOCATE"
    DPallocate.Show vbModal
End Sub

Private Sub mnuShowAll_Click()
    DPallocate.RowSrc = ActRow
    DPallocate.MySource = "SHOW"
    DPallocate.Show vbModal
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    

    With flxDP              '   1       2          3          4       5          6        7         8           9       10              11
        .FormatString = "   |^ DATE|^RECEIPT NO.|<CUSTOMER|>DP AMOUNT|<STATUS|>BALANCE |<INVOICES|<JOB ORDERS |<REMARKS"  'CUSTOMERID   AMTALLOC
        .ColWidth(0) = 300
        For i = 1 To 9
            .ColWidth(i) = (.Width - 300) * 0.12
        Next i
        .ColWidth(3) = (.Width - 300) * 0.2
        .ColWidth(9) = (.Width - 300) * 0.2
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
            .TextMatrix(i, 4) = Format$(.TextMatrix(i, 4), "#,##0.00")
            .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
            If .TextMatrix(i, 5) = "OPEN" Then
                .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
            End If
            If .TextMatrix(i, 5) = "CANCELLED" Then
                .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                .Col = 2
            End If
        Next i
        .ZOrder
    End With
    
    cmdFind.Height = cboStatus.Height
    cmdFind.Top = cboStatus.Top
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

