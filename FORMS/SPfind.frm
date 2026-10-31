VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPfind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Spare Parts"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14070
   Icon            =   "SPfind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   14070
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   180
      Width           =   4455
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   7320
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
      Image           =   "SPfind.frx":058A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptHide 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Hide Option"
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
      Image           =   "SPfind.frx":0B24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptShow 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Show Option"
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
      Image           =   "SPfind.frx":10BE
      cBack           =   -2147483633
   End
   Begin VB.Frame FrameOpt 
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   13695
      Begin VB.ComboBox cboSPType 
         Height          =   315
         Left            =   9840
         TabIndex        =   14
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtModel 
         Height          =   315
         Left            =   1890
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "SPARE PART TYPE:"
         Height          =   255
         Left            =   7680
         TabIndex        =   15
         Top             =   390
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Type Model:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   390
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   45
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      Image           =   "SPfind.frx":1658
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   10920
      TabIndex        =   7
      Top             =   7320
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
      Image           =   "SPfind.frx":1BF2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   12360
      TabIndex        =   8
      Top             =   7320
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
      Image           =   "SPfind.frx":218C
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   5175
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   $"SPfind.frx":2726
   End
   Begin MSFlexGridLib.MSFlexGrid flxBO 
      Height          =   1095
      Left            =   4080
      TabIndex        =   12
      Top             =   9240
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
   Begin lvButton.lvButtons_H cmdBOSQ 
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   7320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "&Back Order && Sales Qty."
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
      Image           =   "SPfind.frx":2806
      cBack           =   -2147483633
   End
   Begin MSComCtl2.UpDown UpDRef 
      Height          =   315
      Left            =   1845
      TabIndex        =   17
      Top             =   180
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Max             =   1
      Enabled         =   -1  'True
   End
   Begin VB.Label lblPartsCatalog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   6960
      Width           =   13575
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   8760
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14160
      m_MDI_Height    =   8235
   End
   Begin VB.Label lblRefSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIND PARTS DESC."
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label lblSqtyBck 
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   7320
      Width           =   5055
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
      Left            =   360
      TabIndex        =   10
      Top             =   8040
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
Attribute VB_Name = "SPfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ShowOpt As Boolean, ComboSrch As Boolean
Dim ActRow As Integer, ActCol As Integer
Public MySource As String
Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub cboSPtype_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If ComboSrch Then
        ComboSrch = False
        Exit Sub
    End If
    With cboSPType
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                ComboSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With

End Sub

Private Sub cboSPtype_Click()
    Call cboSPtype_Change
End Sub

Private Sub cboSPtype_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSPType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSPtype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then ComboSrch = True
End Sub

Private Sub cboSPtype_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub cmdBOSQ_Click()
    Dim MyBO As Integer, MySQ As Integer
    
    With flxSP
        If Val(flxSP.TextMatrix(ActRow, 8)) > 0 Then
            MyBO = GetBackOrder(.TextMatrix(ActRow, 8))
            MySQ = GetSalesQty(.TextMatrix(ActRow, 8))
            lblSqtyBck.Caption = .TextMatrix(ActRow, 2) & ": Back Order=" & MyBO & ", SQ=" & MySQ & " [3 mos.]"
        Else
            lblSqtyBck.Caption = ""
        End If
    End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim rsSP As New ADODB.Recordset
    Dim MySP() As String, Mdl() As String
    Dim myW As String, myW1 As String, myW2 As String, myW3 As String
    Dim MySel As String, i As Integer
    
    If cboFind.Text = "" Then Exit Sub
    MySP = Split(cboFind.Text, " ")
    myW1 = ""
    For i = LBound(MySP) To UBound(MySP)
        If myW1 = "" Then
            'myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            If UpDRef.Value = 0 Then
                myW1 = "(SPAREPARTS.PARTNUMBER) Like '%" & MySP(i) & "%'"
            Else
                myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            End If
            
        Else
            'myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            If UpDRef.Value = 0 Then
                myW1 = myW1 & " AND (SPAREPARTS.PARTNUMBER) Like '%" & MySP(i) & "%'"
            Else
                myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            End If
        End If
    Next i
    myW = myW1
    If ShowOpt Then
        If txtModel.Text <> "" Then
            Mdl = Split(txtModel.Text, " ")
            myW2 = ""
            For i = LBound(Mdl) To UBound(Mdl)
                If myW2 = "" Then
                    myW2 = "(SPAREPARTS.APPLICMODEL) Like '%" & Mdl(i) & "%'"
                Else
                    myW2 = myW2 & " AND (SPAREPARTS.APPLICMODEL) Like '%" & Mdl(i) & "%'"
                End If
            Next i
            myW = "(" & myW & ") AND (" & myW2 & ")"
        End If
        If cboSPType.ListIndex >= 1 Then
            myW = "(" & myW & ") AND ((SPAREPARTS.REMARKS)='" & cboSPType.Text & "')"
        End If
    End If
    '***Access
    'MySel = "SELECT SPAREPARTS.QTYONHAND, SPAREPARTS.PARTNUMBER, [DESCRIPTION] & IIf(IsNull([DESCRIPTION2]),'',' ' & [DESCRIPTION2]) AS MYDESC, SPAREPARTS.APPLICMODEL, SPAREPARTS.STOREPRICE, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.LOCATION, SPAREPARTS.SPAREPARTID, SPAREPARTS.COST From SPAREPARTS WHERE (" & myW & ");"
    
    'mySQL
    MySel = "SELECT SPAREPARTS.QTYONHAND, SPAREPARTS.PARTNUMBER, CONCAT(SPAREPARTS.DESCRIPTION, ' ', If(IsNull(SPAREPARTS.DESCRIPTION2),'', " & _
            "SPAREPARTS.DESCRIPTION2)) AS MYDESC, SPAREPARTS.APPLICMODEL, SPAREPARTS.STOREPRICE, SPAREPARTS.MAXDISCOUNT, " & _
            "SPAREPARTS.LOCATION, SPAREPARTS.SPAREPARTID, SPAREPARTS.COST, SPAREPARTS.DLRPRICE From SPAREPARTS " & _
            "WHERE description not like '%inactive%' and " & myW & ";"

    With rsSP
        .Open MySel, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 1
            Do While Not .EOF
                If MySource = "TRA" Then
                    flxSP.AddItem vbTab & !QtyOnHand & vbTab & !PartNumber & vbTab & !MyDesc & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & Format$(!DlrPrice, "###,##0.00") & vbTab & vbTab & !Location & vbTab & !SparepartID & vbTab & !Cost
                Else
                    flxSP.AddItem vbTab & !QtyOnHand & vbTab & !PartNumber & vbTab & !MyDesc & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & Format$(!STOREPRICE, "###,##0.00") & vbTab & IIf(IsNull(!MaxDiscount) Or !MaxDiscount = 0, "", !MaxDiscount & "%") & vbTab & !Location & vbTab & !SparepartID & vbTab & !Cost
                End If
                .MoveNext
            Loop
            flxSP.Visible = True
'            If ShowOpt Then
''                flxSP.Top = 1680
''                cmdOK.Top = 3480
''                cmdInsert.Top = 3480
''                cmdCancel.Top = 3480
''                lblSqtyBck.Top = 3574
''                cmdBOSQ.Top = 3480
'                flxSP.Top = cmdOptShow.Top + 2
'            Else
'                flxSP.Top = 1080
'                cmdInsert.Top = 2880
'                cmdOK.Top = 2880
'                cmdCancel.Top = 2880
'                lblSqtyBck.Top = 2974
'                cmdBOSQ.Top = 2880
'            End If
            'Me.Height = cmdOK.Top + cmdOK.Height + 580
        Else
            flxSP.Clear
            flxSP.Rows = 2
            'flxSP.FormatString = "      |^Avl.   |<PART NUMBER            |<   DESCRIPTION                         |<APPLIC. MODEL  |>PRICE        |^DISCOUNT       |<LOCATION                  "
            'flxSP.Visible = True
            Call RXResizer1_ResizeComplete
        End If
        .Close
    End With
    flxSP.ColWidth(8) = 0
    flxSP.ColWidth(9) = 0
    Call AddInCombo
    If flxSP.Rows >= 2 Then
        If flxSP.TextMatrix(1, 3) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
    lblSqtyBck.Caption = ""
End Sub

Private Sub AddInCombo()
    Dim X As Integer, w As Boolean
    
    If cboFind.Text = "" Then Exit Sub
    w = False
    For X = 0 To cboFind.ListCount - 1
        If cboFind.Text = cboFind.List(X) Then
            If X = 0 Then
                w = True
                Exit For
            Else
                cboFind.RemoveItem X
                Exit For
            End If
        End If
    Next X
    If w = False Then cboFind.AddItem cboFind.Text, 0
End Sub

Private Sub cmdInsert_Click()
    Dim i As Integer, myDLRPrice As Currency, Qry As String
    
    
    If MySource = "SPSORDER" Then
        With SPSorder.flxOrd
            If .TextMatrix(.Rows - 1, 10) = "" And .TextMatrix(.Rows - 1, 2) = "" Then 'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                     .RemoveItem .Rows - 1
                End If
            End If
            For i = 1 To flxSP.Rows - 1
                myDLRPrice = 0
                If flxSP.TextMatrix(i, 0) <> "" Then
                    Qry = "Select DLRPRICE From SPAREPARTS WHERE SPAREPARTID = " & Val(flxSP.TextMatrix(i, 8)) & " "
                    myDLRPrice = GetFieldValue(Qry)
                    If Not SPinFlx(flxSP.TextMatrix(i, 8)) Then
                        .AddItem vbTab & flxSP.TextMatrix(i, 1) & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 7) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 4) & vbTab & flxSP.TextMatrix(i, 6) & vbTab & PureNum(flxSP.TextMatrix(i, 5)) & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & myDLRPrice
                    End If
                End If
            Next i
            If .Rows = 1 Then .Rows = 2
        End With
    End If
    If MySource = "REQUISITION" Then
        With Requisition.flxOrd
            If .TextMatrix(.Rows - 1, 7) = "" And .TextMatrix(.Rows - 1, 2) = "" Then 'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                     .RemoveItem .Rows - 1
                End If
            End If
            For i = 1 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 8)) Then
                        .AddItem vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 1)
                    End If
                End If
            Next i
            If .Rows = 1 Then .Rows = 2
        End With
    End If
    If MySource = "TRA" Then
        With TRA.flxOrd
            If .TextMatrix(.Rows - 1, 10) = "" And .TextMatrix(.Rows - 1, 2) = "" Then 'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                     .RemoveItem .Rows - 1
                End If
            End If
            For i = 1 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 8)) Then
                        .AddItem vbTab & flxSP.TextMatrix(i, 1) & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 7) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 4) & vbTab & flxSP.TextMatrix(i, 6) & vbTab & PureNum(flxSP.TextMatrix(i, 5)) & vbTab & flxSP.TextMatrix(i, 9)
                    End If
                End If
            Next i
            If .Rows = 1 Then .Rows = 2
        End With
    End If

End Sub

Function SPinFlx(MyID As Long) As Boolean
    Dim i As Integer
    
    If MySource = "SPSORDER" Then
        With SPSorder.flxOrd
            For i = 1 To .Rows - 1
                If MyID = Val(.TextMatrix(i, 10)) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
    If MySource = "REQUISITION" Then
        With Requisition.flxOrd
            For i = 1 To .Rows - 1
                If MyID = .TextMatrix(i, 7) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
    If MySource = "TRA" Then
        With TRA.flxOrd
            For i = 1 To .Rows - 1
                If MyID = Val(.TextMatrix(i, 10)) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
    
End Function

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    'Call SPSorder.ComputeTotal
    Unload Me
End Sub

Private Sub cmdOptHide_Click()
    FrameOpt.Visible = False
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = False
    cmdOptShow.Visible = True
'    flxSP.Top = 1080
'    cmdInsert.Top = 2880
'    cmdOK.Top = 2880
'    cmdCancel.Top = 2880
'    lblSqtyBck.Top = 2974
  '  cmdBOSQ.Top = 2880
    
    flxSP.Top = cmdOptShow.Top + cmdOptShow.Height + 5
    lblPartsCatalog.Top = flxSP.Top + flxSP.Height + 100
    cmdInsert.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    cmdOK.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    cmdCancel.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    lblSqtyBck.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    cmdBOSQ.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    Me.Height = cmdOK.Top + cmdOK.Height + 580
    txtModel.Text = ""
    ShowOpt = False
End Sub

Private Sub cmdOptShow_Click()
    FrameOpt.Visible = True
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = True
    cmdOptShow.Visible = False
    flxSP.Visible = True
'    flxSP.Top = 1680
'    cmdInsert.Top = 3480
'    cmdOK.Top = 3480
'    cmdCancel.Top = 3480
'    lblSqtyBck.Top = 3574
'    cmdBOSQ.Top = 3480
    flxSP.Top = FrameOpt.Top + FrameOpt.Height + 5
    lblPartsCatalog.Top = flxSP.Top + flxSP.Height + 100
    cmdInsert.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    cmdOK.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    cmdCancel.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    lblSqtyBck.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    cmdBOSQ.Top = lblPartsCatalog.Top + lblPartsCatalog.Height + 100
    Me.Height = cmdOK.Top + cmdOK.Height + 580
    ShowOpt = True
End Sub

Private Sub flxSP_Click()
If flxSP.TextMatrix(ActRow, 2) <> "" Then
    Call getCatalogPath(flxSP.TextMatrix(flxSP.Row, 2))
End If
End Sub

Public Sub getCatalogPath(myPartNum As String)
    If ConnectPartsCatalog = True Then
        Qry = "call partscatalog.getCatalogPath('" & myPartNum & "') "
        With rsAnyUse
            If .State = adStateOpen Then .Close
            .Open Qry, conPartsCat, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                lblPartsCatalog.Caption = ""
                lblPartsCatalog.Caption = "CATALOG PATH: " & IIf(IsNull(!MOTORCYCLE), "", !MOTORCYCLE)
            Else
                lblPartsCatalog.Caption = ""
            End If
            .Close
        End With
    Else
        MsgBox "Cannot Connect To Parts Catalog Database!", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
End Sub

Private Sub flxSP_DblClick()
        If flxSP.TextMatrix(flxSP.Row, 0) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
End Sub

Private Sub flxSP_EnterCell()
    ActRow = flxSP.Row
    If flxSP.TextMatrix(ActRow, 2) <> "" Then
        Call getCatalogPath(flxSP.TextMatrix(flxSP.Row, 2))
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

Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Val(flxSP.TextMatrix(flxSP.Row, 1)) = 0 Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuSP)
    End If
End Sub

Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xRow As Integer, xBal As Integer, i As Integer
    
    'Clear FlxStock
    '***Access
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyId & ")) ORDER BY SPITR.TRANSDATE;"
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.`ORDER` From SPITR " & _
          "WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.SPAREPARTID=" & MyID & " ORDER BY SPITR.TRANSDATE;"
    
    
    flxBO.Rows = 1
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxBO.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !Order
                .MoveNext
            Loop
        End If
        .Close
    End With
    xBal = 0
    With flxBO
        For i = 1 To .Rows - 1
            'Display PO in Stock Card and obtain Balance
            xBal = xBal + Val(.TextMatrix(i, 2))
            'flxStock.AddItem .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 1) & vbTab & vbTab & .TextMatrix(i, 2) & vbTab & vbTab & xBal   '.TextMatrix(i, 2)
            'Acquire Packing list for P.O.
            MyPO = .TextMatrix(i, 1)
            '***Access
            'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyId & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
            'mySQL
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyPO & "' ORDER BY SPITR.TRANSDATE;"
            
            rsRec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRec.RecordCount > 0 Then
                rsRec.MoveFirst
                Do While Not rsRec.EOF
                    xBal = xBal - rsRec!RECEIVE
                    'flxStock.AddItem Format(rsRec!TransDate, "mm/dd/yyyy") & vbTab & MyPO & vbTab & rsRec!Reference & vbTab & vbTab & rsRec!Receive & vbTab & xBal
                    rsRec.MoveNext
                Loop
            End If
            rsRec.Close
        Next i
    End With
    GetBackOrder = xBal
End Function

Function GetSalesQty(MyID As Long) As Integer
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String, StartDate As Date, EndDate As Date
    Dim MySales As Integer, MyRet As Integer
    
    StartDate = DateSerial(Year(DateAdd("m", -3, Date)), Month(DateAdd("m", -3, Date)), 1)
    EndDate = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), Day(DateAdd("d", -1, DateSerial(Year(Date), Month(Date), 1))))
    '***Access
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(StartDate, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(EndDate, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.SPAREPARTID=" & MyID & ";"
    
    With rsQty
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    '***Access
    'Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyId & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & StartDate & "# And (SPITR.TRANSDATE)<=#" & EndDate & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(StartDate, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(EndDate, "yyyy-mm-dd") & "';"
    
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function


Private Sub Form_Load()
    Dim rsSPType As New ADODB.Recordset
    Dim Qry As String
    
    flxSP.ColWidth(8) = 0
    Call cmdOptHide_Click
    If MySource = "SPSORDER" Then
        If SPSorder.JustScroll = "True" Then
            cmdInsert.Enabled = False
            cmdOK.Enabled = False
        End If
    End If
    If MySource = "REQUISITION" Then
        If Requisition.JustScroll = "True" Then
            cmdInsert.Enabled = False
            cmdOK.Enabled = False
        End If
    End If
    If MySource = "TRA" Then
        If TRA.JustScroll = "True" Then
            cmdInsert.Enabled = False
            cmdOK.Enabled = False
        End If
    End If
    
    Qry = "SELECT DISTINCT SUPPLIER.SPTYPE From SUPPLIER ORDER BY SUPPLIER.SPTYPE;"
    With rsSPType
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            cboSPType.Clear
            cboSPType.AddItem "ALL TYPE"
            .MoveFirst
            Do While Not .EOF
                cboSPType.AddItem IIf(IsNull(!SPtype), "", !SPtype)
                .MoveNext
            Loop
        End If
        .Close
    End With
    cboSPType.ListIndex = 0
    
    UpDRef.Value = 1
    'LYRmain.sBar.Panels(6).Text = "Find Spare Parts"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'LYRmain.sBar.Panels(6).Text = ""
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

Private Sub RXResizer1_ResizeComplete()

    With flxSP
        .FormatString = "      |^Avl.   |<PART NUMBER            |<   DESCRIPTION                         |<APPLIC. MODEL  |>PRICE        |^DISCOUNT       |<LOCATION                  "
        .ColWidth(0) = 300
        .ColWidth(1) = .Width * 0.05
        .ColWidth(2) = .Width * 0.18
        .ColWidth(3) = .Width * 0.25
        .ColWidth(4) = .Width * 0.15
        .ColWidth(5) = .Width * 0.08
        .ColWidth(6) = .Width * 0.1
        .ColWidth(7) = .Width * 0.12
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        
    End With
End Sub

Private Sub UpDRef_Change()
    If UpDRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    Else
        lblRefSearch.Caption = "FIND PARTS DESC."
    End If
End Sub
