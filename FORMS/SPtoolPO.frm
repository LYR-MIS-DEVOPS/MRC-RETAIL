VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SPToolPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Top Selling Spare Parts"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12315
   Icon            =   "SPtoolPO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   4695
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8281
      _Version        =   393216
      RowHeightMin    =   285
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   11775
      Begin VB.TextBox txtItemType 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtDateC 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtNumSP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboPeriod 
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   225
         Width           =   3975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ITEM TYPE:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   615
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE CREATED:"
         Height          =   255
         Left            =   7080
         TabIndex        =   12
         Top             =   615
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "NO. OF SPAREPARTS:"
         Height          =   255
         Left            =   7200
         TabIndex        =   10
         Top             =   255
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SELECT SALES PERIOD:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   255
         Width           =   1935
      End
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSuppCode 
      Height          =   285
      Left            =   -120
      TabIndex        =   4
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Insert"
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
      Image           =   "SPtoolPO.frx":058A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   6240
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
      Image           =   "SPtoolPO.frx":0B24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   10800
      TabIndex        =   2
      Top             =   6240
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
      Image           =   "SPtoolPO.frx":10BE
      cBack           =   -2147483633
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
      Left            =   0
      TabIndex        =   3
      Top             =   6360
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
Attribute VB_Name = "SPToolPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActRow As Integer, ActCol As Integer


Private Sub cboPeriod_Click()
    Dim rsSal As New ADODB.Recordset
    Dim Qry As String, MyCov As String, i As Integer, MonRpt As Integer, j As Integer
    
    If cboPeriod.ListIndex = -1 Then Exit Sub
    Qry = "SELECT TOPSALESUMM.* From TOPSALESUMM WHERE (((TOPSALESUMM.TOPSALESID)=" & cboPeriod.ItemData(cboPeriod.ListIndex) & "));"
    With rsSal
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtDateC.Text = Format$(!DateCreated, "mm/dd/yyyy")
            txtItemType.Text = !ItemType
            txtNumSP.Text = "TOP " & !TopValue & " only"
            MonRpt = IIf(IsNull(!MonthInRpt), 0, !MonthInRpt)
        End If
        .Close
    End With
    'Build Qry to Display in FlxSP
    If MonRpt > 0 Then
        MyCov = ""
        For i = 1 To MonRpt
            MyCov = MyCov & " TOPSALEDTLS.MONTH" & i & ","
        Next i
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.QTYONHAND, " & MyCov & " TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((TOPSALEDTLS.TOPSALESID) = " & Val(cboPeriod.ItemData(cboPeriod.ListIndex)) & ")) ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
    Else
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((TOPSALEDTLS.TOPSALESID) = " & Val(txtTopSalesID.Text) & ")) ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
    End If
    'txtSQL.Text = Qry
    'CboSort1.ListIndex = 7
    'CboAD1.ListIndex = 0
    flxSP.Clear
    If rsSal.State = adStateOpen Then rsSal.Close
    rsSal.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSP.DataSource = rsSal
    'Arrange FlxSP
    flxSP.ColWidth(0) = 300
    flxSP.ColWidth(1) = 0
    flxSP.ColWidth(2) = (flxSP.Width - 300) * 0.14
    flxSP.ColWidth(3) = (flxSP.Width - 300) * 0.22
    flxSP.ColWidth(flxSP.Cols - 1) = 0
    For i = 4 To flxSP.Cols - 1
        flxSP.ColWidth(i) = (flxSP.Width - 300) * 0.05
    Next i
    flxSP.TextMatrix(0, 4) = "QH"
    flxSP.TextMatrix(0, 5) = "BO"
    flxSP.TextMatrix(0, 6) = "PO"
    If MonRpt > 0 Then
        j = MonRpt
        For i = 7 To 6 + MonRpt
            flxSP.TextMatrix(0, i) = UCase$(Format$(DateAdd("m", j * -1, Date), "MMMYY"))
            j = j - 1
        Next i
    End If
    flxSP.TextMatrix(0, flxSP.Cols - 4) = "AVG."
    flxSP.TextMatrix(0, flxSP.Cols - 3) = "TTL."
    flxSP.TextMatrix(0, flxSP.Cols - 2) = "COST"
    flxSP.TextMatrix(0, flxSP.Cols - 1) = "RANK"
    'Format Records
    Pbar.Value = 0
    Pbar.Max = flxSP.Rows - 1
    Pbar.Visible = True
    With flxSP
        For i = 1 To .Rows - 1
            .TextMatrix(i, 6) = ""  'Format(.TextMatrix(i, .Cols - 4), "#,##0.00")
            .TextMatrix(i, .Cols - 2) = Format(.TextMatrix(i, .Cols - 2), "#,##0.00")
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    Pbar.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    Dim i As Integer
        
    With SPordering.flxParts
        If .TextMatrix(.Rows - 1, 1) = "" Then  'Remove last row if empty
            If .Rows = 2 Then
                .Rows = 1
            Else
                .RemoveItem .Rows - 1
            End If
        End If
        For i = 1 To flxSP.Rows - 1
            If flxSP.TextMatrix(i, 0) <> "" Then
                If Not SPinFlx(flxSP.TextMatrix(i, 1)) Then 'x
                    '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & flxSP.TextMatrix(i, 6) & vbTab & vbTab & flxSP.TextMatrix(i, 8)
                    .AddItem .Rows & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 6) & vbTab & flxSP.TextMatrix(i, flxSP.Cols - 2) & vbTab & vbTab & flxSP.TextMatrix(i, 1)
                End If
            End If
        Next i
        If .Rows = 1 Then .AddItem "1"
    End With
End Sub

Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    With SPordering.flxParts
        For i = 1 To .Rows - 1
            If myPN = .TextMatrix(i, 1) Then
                SPinFlx = True
                Exit Function
            End If
        Next i
    End With
    SPinFlx = False
End Function

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Call SPordering.ComputeTotal
    Unload Me
End Sub


Private Sub flxSP_Click()
    ActRow = flxSP.Row
    ActCol = flxSP.Col
    If ActRow = 0 Or ActCol <> 6 Then
        txtOrder.Visible = False
        Exit Sub
    End If
    With txtOrder
        .Top = flxSP.CellTop + flxSP.Top
        .Left = flxSP.CellLeft + flxSP.Left
        .Width = flxSP.CellWidth
        .Text = Trim(flxSP.Text)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With
End Sub

Private Sub flxSP_DblClick()
        If flxSP.TextMatrix(flxSP.Row, 0) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
End Sub

Private Sub flxSP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxSP.TextMatrix(flxSP.Row, 0) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
    End If
End Sub

Private Sub flxSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSP_Click
End Sub

Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Val(flxSP.TextMatrix(flxSP.Row, 1)) = 0 Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuSP)
    End If
End Sub

Private Sub flxSP_Scroll()
    txtOrder.Visible = False
End Sub

Private Sub Form_Load()
    Dim rsTOP As New ADODB.Recordset
    Dim Qry As String, j As Integer, tmpDate As Date
    
    Qry = "SELECT TOPSALESUMM.DATECREATED, TOPSALESUMM.SALESPERIOD, TOPSALESUMM.ITEMTYPE, TOPSALESUMM.TOPVALUE, TOPSALESUMM.NOTE, TOPSALESUMM.TOPSALESID " & _
          "From TOPSALESUMM ORDER BY TOPSALESUMM.DATECREATED DESC;"
    With rsTOP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            cboPeriod.Clear
            .MoveFirst
            Do While Not .EOF
                j = IIf(IsNull(!SalesPeriod), 0, !SalesPeriod)
                If j = 0 Then   'From Beginning: Get Date difference 1/1/2009
                    cboPeriod.AddItem "JAN. '08 TO " & Format(!DateCreated, "MMM. 'YY")
                Else
                    cboPeriod.AddItem UCase$(Format$(DateAdd("M", j * -1, !DateCreated), "MMM. 'YY")) & " TO " & UCase$(Format$(DateAdd("M", -1, !DateCreated), "MMM. 'YY"))
                End If
                cboPeriod.ItemData(cboPeriod.ListCount - 1) = !TopSalesID
                .MoveNext
            Loop
        End If
        .Close
    End With
    With flxSP
                        '  1       2             3         4   5            6                                                            7                                                       8                                                9       10      11    12
        .FormatString = "   |  |< PART NUMBER  |<DESCRIPTION|^QH|^BO|^" & UCase$(Format(DateAdd("M", -3, Date), "MMM. 'YY")) & "|^" & UCase$(Format(DateAdd("M", -2, Date), "MMM. 'YY")) & "|^" & UCase$(Format(DateAdd("M", -1, Date), "MMM. 'YY")) & "|^AVG. MO|^TTL SALES|^COST|^RANK"
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        For i = 2 To 12
            .ColWidth(i) = (.Width - 300) * 0.06
        Next i
        .ColWidth(2) = (.Width - 300) * 0.14
        .ColWidth(3) = (.Width - 300) * 0.22
        .ZOrder
    End With
    With SPordering
        If .LockFlx = "True" Then
            cmdInsert.Enabled = False
            cmdOK.Enabled = False
        End If
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'LYRmain.sBar.Panels(6).Text = ""
    'If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub


Private Sub mnuMark_Click()
    With flxSP
        .Col = 0
        .CellFontName = "WINGDINGS"
        .CellFontSize = 12
        .Text = lblChk.Caption
        .CellAlignment = 4
    End With
End Sub

Private Sub mnuMarkAll_Click()
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

Private Sub mnuUnmark_Click()
    flxSP.TextMatrix(flxSP.Row, 0) = ""
End Sub

Private Sub mnuUnmarkAll_Click()
    Dim i As Integer
    
    For i = 1 To flxSP.Rows - 1
        flxSP.TextMatrix(i, 0) = ""
    Next i
End Sub

Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtOrder.Text = flxSP.TextMatrix(ActRow, 6)
        flxSP.SetFocus
        flxSP.Row = ActRow
        flxSP.Col = 6
        txtOrder.Visible = False
    End If
End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxSP.SetFocus
        flxSP.Row = IIf(ActRow + 1 <= flxSP.Rows - 1, ActRow + 1, ActRow)
        flxSP.Col = 6
    End If
End Sub

Private Sub txtOrder_LostFocus()
    
    With flxSP
        .TextMatrix(ActRow, 6) = txtOrder.Text
        txtOrder.Visible = False
        If Val(.TextMatrix(ActRow, 6)) > 0 Then
            .Row = ActRow
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .Col = 6
        Else
            .TextMatrix(ActRow, 6) = ""
            .TextMatrix(ActRow, 0) = ""
            .Col = 6
        End If
    End With
End Sub
