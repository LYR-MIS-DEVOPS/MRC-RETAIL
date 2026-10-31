VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SalesRetInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice for Return"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   Icon            =   "SalesRetInv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdChkAll 
      Height          =   300
      Left            =   200
      TabIndex        =   5
      Top             =   200
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   529
      Caption         =   "ü"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "Qty. Purchased: 0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&OK"
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
      Image           =   "SalesRetInv.frx":038A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
      Image           =   "SalesRetInv.frx":0924
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   "      |^QTY.          |^UNIT          |<PART NUMBER                     |<   DESCRIPTION                        "
   End
   Begin VB.Label lblX 
      Caption         =   "û"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblChk 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuSP 
      Caption         =   "SP"
      Visible         =   0   'False
      Begin VB.Menu mnuMark 
         Caption         =   "Mark"
      End
      Begin VB.Menu mnuUnmark 
         Caption         =   "Unmark"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkAll 
         Caption         =   "Mark All"
      End
      Begin VB.Menu mnuUnmarkAll 
         Caption         =   "UnMark All"
      End
   End
End
Attribute VB_Name = "SalesRetInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShowOpt As Boolean
Dim ActRow As Integer, ActCol As Integer


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChkAll_Click()
    Dim i As Integer
    
    With flxSP
        For i = 1 To .Rows - 1
            .Row = i
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .TextMatrix(i, 0) = lblChk.Caption
        Next i
    End With
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, Ans As Integer
    
    With SalesReturn.flxParts
        If .TextMatrix(.Rows - 1, 7) <> "" And .Rows >= 2 Then 'Prompt user that previous entry will be removed
            Ans = MsgBox("All your previous entries will be removed!" & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, "Sales Return")
            If Ans = vbNo Then Exit Sub
        End If
        .Rows = 1
        For i = 1 To flxSP.Rows - 1
            If flxSP.TextMatrix(i, 0) <> "" Then
                .AddItem vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 4) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & flxSP.TextMatrix(i, 6) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & flxSP.TextMatrix(i, 8)
            End If
        Next i
    End With
    Call SalesReturn.ComputeTotal
    Unload Me
End Sub


Private Sub flxSP_Click()
    ActRow = flxSP.Row
    ActCol = flxSP.Col
    
    If ActCol > 1 Or ActRow = 0 Then Exit Sub
    
    Select Case ActCol
        Case 0
            If flxSP.TextMatrix(ActRow, 0) = "" And Val(flxSP.TextMatrix(ActRow, 1)) > 0 Then
                Call mnuMark_Click
            Else
                flxSP.TextMatrix(ActRow, 0) = ""
            End If
        Case 1
            With txtFlx
                .Top = flxSP.CellTop + flxSP.Top
                .Left = flxSP.CellLeft + flxSP.Left
                .Width = flxSP.CellWidth
                .Text = flxSP.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
            Call QtyTip(True)
    End Select
End Sub

Private Sub QtyTip(Disp As Boolean)
    If Disp Then
        With txtQtyTip
            .Top = flxSP.CellTop + flxSP.Top + 290
            .Left = flxSP.Left + flxSP.ColWidth(0) + 300
            .Text = "Qty. Purchased: " & flxSP.TextMatrix(ActRow, 8)
            .Visible = True
            .ZOrder
        End With
    Else
        txtQtyTip.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, MyRef As String
    Dim MyQty As Integer
    
    '***Access
    'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [PRICE]*[OUT] AS TTLAMT, SPITR.SPAREPARTID, SPITR.REFERENCE " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='SI#-" & SalesReturn.txtInvNo.Text & "'));"
    'mySQl
    
    If InStr(UCase$(SalesReturn.txtInvNo.Text), "TRA#:") > 0 Then
        MyRef = SalesReturn.txtInvNo.Text
    Else
        MyRef = "SI#-" & SalesReturn.txtInvNo.Text
    End If
    Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.OUT AS TTLAMT, SPITR.SPAREPARTID, SPITR.REFERENCE " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "' AND SPITR.REMARKS = '" & IIf(SalesReturn.UpDnINV.Value = 1, "CASH", "CREDIT") & "';"
    
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 1 Then
            .MoveFirst
            flxSP.Rows = 1
            Do While Not .EOF
                If InStr(!Description, "ENGINE OIL") > 0 Or InStr(!Description, "4T OIL") > 0 Then
                    MyUnit = IIf(!OUT > 1, "QRTS.", "QRT.")
                Else
                    MyUnit = IIf(!OUT > 1, "PCS.", "PC.")
                End If
                MyQty = !OUT - ReturnedQty(!SPAREPARTID)
                If MyQty > 0 Then
                    flxSP.AddItem vbTab & MyQty & vbTab & MyUnit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "###,##0.00") & vbTab & Format(!TtlAmt, "###,##0.00") & vbTab & !SPAREPARTID & vbTab & MyQty
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    flxSP.Col = 0
    flxSP.Row = 0
    cmdChkAll.Top = flxSP.CellTop + flxSP.Top
    cmdChkAll.Left = flxSP.CellLeft + flxSP.Left
    MotoRep.sBar.Panels(6).Text = "Select Item in Invoice"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = "Sales Return"
End Sub

Function ReturnedQty(SPid As Long) As Integer
    Dim rsRetQ As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.SPAREPARTID)=" & SPid & ") AND ((SPITR.PACKINGLIST)='" & SalesReturn.txtInvNo.Text & "'));"
    'mySQL
    Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.SPAREPARTID=" & SPid & " AND SPITR.PACKINGLIST='" & SalesReturn.txtInvNo.Text & "';"
    
    With rsRetQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ReturnedQty = !RECEIVE
        Else
            ReturnedQty = 0
        End If
        .Close
    End With
End Function

Private Sub mnuMark_Click()
    With flxSP
        .Col = 0
        .CellFontName = "WINGDINGS"
        .CellFontSize = 12
        .Text = lblChk.Caption
        .CellAlignment = 4
    End With
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then Call txtFlx_LostFocus
End Sub

Private Sub txtFlx_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtFlx.Tag = "skiplostfocus"
        txtFlx.Visible = False
        flxSP.SetFocus
        flxSP.Col = ActCol
    End If
End Sub

Private Sub txtFlx_LostFocus()
    Dim Price As Single
        
    If txtFlx.Tag = "skiplostfocus" Then
        txtFlx.Tag = ""
        Exit Sub
    End If
    If Val(txtFlx.Text) > Val(flxSP.TextMatrix(ActRow, 8)) Or Val(txtFlx.Text) = 0 Then
        MsgBox "Qty. Return must not be Zero or Greater than Qty. Purchased!", vbExclamation + vbOKOnly, "Validation Error"
        txtFlx.SelStart = 0
        txtFlx.SelLength = Len(txtFlx.Text)
        txtFlx.SetFocus
    Else
        flxSP.TextMatrix(ActRow, 1) = txtFlx.Text
        txtFlx.Visible = False
        If flxSP.TextMatrix(ActRow, 5) <> "" Then
            Price = PureNum(flxSP.TextMatrix(ActRow, 5))
            flxSP.TextMatrix(ActRow, 6) = Format$(Val(txtFlx.Text) * Price, "#,##0.00")
    '        Call ComputeTotal
        End If
        Call QtyTip(False)
        flxSP.SetFocus
        flxSP.Col = 1
    End If
End Sub

