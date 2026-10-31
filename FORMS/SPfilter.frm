VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPfilter 
   Caption         =   "Criteria"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   Icon            =   "SPfilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNoFactor 
      Caption         =   "Do not Include Parts with existing Factor"
      Height          =   255
      Left            =   1200
      TabIndex        =   36
      Top             =   70
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox cboSupplier 
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   29
      Top             =   4320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cboSupplier 
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtSupCode 
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   34
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSupCode 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   33
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSupCode 
      Height          =   285
      Index           =   0
      Left            =   3720
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cboSupplier 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   4680
      TabIndex        =   28
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "OK"
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
      Image           =   "SPfilter.frx":058A
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   1200
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton optOr2 
         Caption         =   "Or )"
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         Top             =   120
         Width           =   615
      End
      Begin VB.OptionButton optAnd2 
         Caption         =   "And"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   150
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "( Operator :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1200
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton optOr1 
         Caption         =   "Or )"
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
      Begin VB.OptionButton optAnd1 
         Caption         =   "And"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   150
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "( Operator :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.ComboBox cboMatch 
      Height          =   315
      Index           =   2
      ItemData        =   "SPfilter.frx":0B24
      Left            =   1200
      List            =   "SPfilter.frx":0B31
      TabIndex        =   19
      Text            =   "Any Part of Field"
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   17
      Top             =   4320
      Width           =   3135
   End
   Begin VB.ComboBox cboLookIn 
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   15
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton cmdMore2 
      Caption         =   "More>>"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   855
   End
   Begin VB.ComboBox cboMatch 
      Height          =   315
      Index           =   1
      ItemData        =   "SPfilter.frx":0B65
      Left            =   1200
      List            =   "SPfilter.frx":0B72
      TabIndex        =   12
      Text            =   "Any Part of Field"
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   10
      Top             =   2520
      Width           =   3135
   End
   Begin VB.ComboBox cboLookIn 
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton cmdMore1 
      Caption         =   "More>>"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1620
      Width           =   855
   End
   Begin VB.ComboBox cboMatch 
      Height          =   315
      Index           =   0
      ItemData        =   "SPfilter.frx":0BA6
      Left            =   1200
      List            =   "SPfilter.frx":0BB3
      TabIndex        =   5
      Text            =   "Any Part of Field"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.ComboBox cboLookIn 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   4680
      TabIndex        =   30
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Cancel"
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
      Image           =   "SPfilter.frx":0BE7
      cBack           =   -2147483633
   End
   Begin VB.Label Label10 
      Caption         =   "Match :"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Find What :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Look in"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Match :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Find What :"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Look in"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Match :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Value"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Look in"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "SPfilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fldNdx As String
Public MySource As String
Public fldNames As String
Public fldDates As String
Public fldSupplier As Integer
Dim CurrentH As Integer, SuppSrch As Boolean
Dim rsSupplier As New ADODB.Recordset
   
Private Sub cboLookIn_Change(Index As Integer)
    If cboLookIn(Index).ListIndex = fldSupplier Then
        cboSupplier(Index).Visible = True
        txtFind(Index).Visible = False
    Else
        cboSupplier(Index).Visible = False
        txtFind(Index).Visible = True
    End If
End Sub

Private Sub cboLookIn_Click(Index As Integer)
    Dim y As String
    y = Trim(Str(cboLookIn(Index).ListIndex))
    If InStr(fldNdx, y) = 0 And InStr(fldDates, y) = 0 Then
        With cboMatch(Index)
            .Clear
            .AddItem "Any Part of Field"
            .AddItem "Whole Field"
            .AddItem "Start of Field"
            .Text = "Any Part of Field"
        End With
    Else
        With cboMatch(Index)
            .Clear
            .AddItem "Equal"
            .AddItem "Less Than"
            .AddItem "Greater Than"
            .AddItem "Less Than Equal"
            .AddItem "Greater Than Equal"
            .Text = "Equal"
        End With
    End If
    If cboLookIn(Index).ListIndex = fldSupplier Then
        cboSupplier(Index).Visible = True
        txtFind(Index).Visible = False
        cboMatch(Index).ListIndex = 1
        cboMatch(Index).Locked = True
    Else
        cboSupplier(Index).Visible = False
        txtFind(Index).Visible = True
        cboMatch(Index).ListIndex = 0
        cboMatch(Index).Locked = False
    End If
End Sub

Private Sub cboSupplier_Change(Index As Integer)
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboSupplier(Index)
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsSupplier.MoveFirst
            rsSupplier.Move i
            txtSupCode(Index).Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            txtFind(Index).Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSupCode(Index).Text = ""
            txtFind(Index).Text = ""
        End If
    End With
End Sub

Private Sub cboSupplier_Click(Index As Integer)
    Call cboSupplier_Change(Index)
End Sub

Private Sub cboSupplier_GotFocus(Index As Integer)
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSupplier(Index).hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdMore1_Click()
    If cmdMore1.Caption = "More>>" Then
        cmdMore1.Caption = "<<Less"
        CurrentH = 4350
        Me.Height = 4350
        Frame1.Visible = True
    Else
        cmdMore1.Caption = "More>>"
        CurrentH = 2520
        Me.Height = 2520
        Frame1.Visible = False
    End If
End Sub

Private Sub cmdMore2_Click()
    If cmdMore2.Caption = "More>>" Then
        cmdMore2.Caption = "<<Less"
        cmdMore1.Enabled = False
        CurrentH = 5505
        Me.Height = 5505
        Frame2.Visible = True
    Else
        cmdMore2.Caption = "More>>"
        cmdMore1.Enabled = True
        CurrentH = 4350
        Me.Height = 4350
        Frame2.Visible = False
    End If
End Sub

Private Sub cmdOK_Click()
    Dim Criteria(3) As String, Qctr As String, GrpCrit As String, Oprtr As String
    Dim Arrfld As Variant, xCnt As Integer
    Dim Oprtr1 As String, Oprtr2 As String, RemOpr(3) As String
    
    Arrfld = Split(fldNames, "*")
    xCnt = 0
    If cmdMore1.Caption = "<<Less" Then xCnt = 1
    If cmdMore2.Caption = "<<Less" Then xCnt = 2
'VALIDATION
    msg = ""
    For X = 0 To xCnt
        If cboLookIn(X).Text = "" Then msg = "Incomplete Search Entry!"
        If txtFind(X).Text = "" Then msg = "Incomplete Search Entry!"
    Next X
    If msg <> "" Then
        MsgBox msg, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    For X = 0 To xCnt
        If UCase(cboMatch(X).Text) = "ANY PART OF FIELD" Then
            'Remove parenthesis
            Criteria(X) = ") Like '%" & txtFind(X).Text & "%'"
            
            'Criteria(X) = " Like '%" & txtFind(X).Text & "%'"
            RemOpr(X) = " Like "
        ElseIf UCase(cboMatch(X).Text) = "START OF FIELD" Then
            'Remove parenthesis
            Criteria(X) = ") Like '" & txtFind(X).Text & "%'"
            
            'Criteria(X) = " Like '" & txtFind(X).Text & "%'"
            
            RemOpr(X) = " Starts with "
        ElseIf UCase(cboMatch(X).Text) = "WHOLE FIELD" Then
            'Remove parenthesis
            Criteria(X) = ") Like '" & txtFind(X).Text & "'"
            'Criteria(X) = " Like '" & txtFind(X).Text & "'"
            
            RemOpr(X) = " Equal to "
        ElseIf UCase(cboMatch(X).Text) = "EQUAL" And InStr(fldDates, cboLookIn(X).ListIndex) <> 0 Then
            'Criteria(X) = ") = #" & txtFind(X).Text & "#"
            'mySQL
            'Criteria(X) = " = '" & Format(txtFind(X).Text, "yyyy-mm-dd") & "'"
            Criteria(X) = ") = '" & Format(txtFind(X).Text, "yyyy-mm-dd") & "'"
            
            RemOpr(X) = " = "
        ElseIf UCase(cboMatch(X).Text) = "LESS THAN" And InStr(fldDates, cboLookIn(X).ListIndex) <> 0 Then
            'Criteria(X) = ") < #" & txtFind(X).Text & "#"
            'mySQL
            'Criteria(X) = " < '" & Format(txtFind(X).Text, "yyyy-mm-dd") & "'"
            Criteria(X) = ") < '" & Format(txtFind(X).Text, "yyyy-mm-dd") & "'"
            
            RemOpr(X) = " < "
        ElseIf UCase(cboMatch(X).Text) = "GREATER THAN" And InStr(fldDates, cboLookIn(X).ListIndex) <> 0 Then
            'Criteria(X) = ") > #" & txtFind(X).Text & "#"
            'mySQL
            'Criteria(X) = " > '" & Format(txtFind(X).Text, "yyyy-mm-dd") & "'"
            Criteria(X) = " > '" & Format(txtFind(X).Text, "yyyy-mm-dd") & "'"
            
            RemOpr(X) = " > "
        ElseIf UCase(cboMatch(X).Text) = "LESS THAN EQUAL" And InStr(fldDates, cboLookIn(X).ListIndex) <> 0 Then
            'Criteria(X) = ") <= #" & txtFind(X).Text & "#"
            'mySQL
            Criteria(X) = ") <= '" & Format(txtFind(X).Text, "yyyy-mm-dd") & "'"
            
            RemOpr(X) = " <= "
        ElseIf UCase(cboMatch(X).Text) = "GREATER THAN EQUAL" And InStr(fldDates, cboLookIn(X).ListIndex) <> 0 Then
            'Criteria(X) = ") >= #" & txtFind(X).Text & "#"
            'mySQL
            Criteria(X) = " >= '" & Format(txtFind(X).Text, "yyyy-mm-dd") & "'"
            RemOpr(X) = " >= "
        ElseIf UCase(cboMatch(X).Text) = "EQUAL" Then
            'Criteria(X) = ") = " & Val(txtFind(X).Text)
            
            Criteria(X) = ") = " & Val(txtFind(X).Text)
            RemOpr(X) = " = "
        ElseIf UCase(cboMatch(X).Text) = "LESS THAN" Then
            Criteria(X) = " < " & Val(txtFind(X).Text)
            RemOpr(X) = " < "
        ElseIf UCase(cboMatch(X).Text) = "GREATER THAN" Then
            Criteria(X) = " > " & Val(txtFind(X).Text)
            RemOpr(X) = " > "
        ElseIf UCase(cboMatch(X).Text) = "LESS THAN EQUAL" Then
            Criteria(X) = " <= " & Val(txtFind(X).Text)
            RemOpr(X) = " <= "
        ElseIf UCase(cboMatch(X).Text) = "GREATER THAN EQUAL" Then
            Criteria(X) = " >= " & Val(txtFind(X).Text)
            RemOpr(X) = " >= "
        End If
    Next X
    
    Select Case xCnt
    Case 0
        'Remove ( )
        'Qctr = "Where (((" & Arrfld(cboLookIn(0).ListIndex) & Criteria(0) & "))"
        'GrpCrit = cboLookIn(0).Text & RemOpr(0) & IIf(txtSupCode(0).Text <> "", cboSupplier(0).Text, txtFind(0).Text)
        
        Qctr = "Where (((" & Arrfld(cboLookIn(0).ListIndex) & Criteria(0) & "))"
        GrpCrit = cboLookIn(0).Text & RemOpr(0) & IIf(txtSupCode(0).Text <> "", cboSupplier(0).Text, txtFind(0).Text)
        Case 1
       Oprtr = IIf(optAnd1.Value = True, " And ", " Or ")
       
       Qctr = "Where (((" & Arrfld(cboLookIn(0).ListIndex) & Criteria(0) & ") " & _
       Oprtr & "((" & Arrfld(cboLookIn(1).ListIndex) & Criteria(1) & "))"
       GrpCrit = cboLookIn(0).Text & RemOpr(0) & IIf(txtSupCode(0).Text <> "", cboSupplier(0).Text, txtFind(0).Text) & Oprtr & _
       cboLookIn(1).Text & RemOpr(1) & IIf(txtSupCode(1).Text <> "", cboSupplier(1).Text, txtFind(1).Text)
        
    Case 2
       Oprtr1 = IIf(optAnd1.Value = True, " And ", " Or ")
       Oprtr2 = IIf(optAnd2.Value = True, " And ", " Or ")
       
       Qctr = "Where " & Arrfld(cboLookIn(0).ListIndex) & Criteria(0) & " " & _
       Oprtr1 & "((" & Arrfld(cboLookIn(1).ListIndex) & Criteria(1) & ") " & _
       Oprtr2 & "((" & Arrfld(cboLookIn(2).ListIndex) & Criteria(2) & "))"
       GrpCrit = cboLookIn(0).Text & RemOpr(0) & IIf(txtSupCode(0).Text <> "", cboSupplier(0).Text, txtFind(0).Text) & Oprtr1 & _
       cboLookIn(1).Text & RemOpr(1) & IIf(txtSupCode(1).Text <> "", cboSupplier(1).Text, txtFind(1).Text) & Oprtr2 & _
       cboLookIn(2).Text & RemOpr(2) & IIf(txtSupCode(2).Text <> "", cboSupplier(2).Text, txtFind(2).Text)
    End Select
    
Select Case MySource
        Case "DETAILEDRECEIVING"
            With InvDtlRecRpt
                .txtCriteria.Text = Left$(Qctr, Len(Qctr) - 1) & " And ((SPITR.TRANSDATE)>='" & Format(.dtpFrom.Value, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(.dtpTo.Value, "yyyy-mm-dd") & "') AND ((SPITR.TRANSTYPE)='RECEIVE'))"
                .lblRemarks.Caption = "Remarks: " & GrpCrit
                Call .ShowFilteredRecords
            End With
        Case "DETAILEDSALES"
            With InvDtlSalRpt
                .txtCriteria.Text = Left$(Qctr, Len(Qctr) - 1) & " And ((SPITR.TRANSDATE)>='" & Format(.dtpFrom.Value, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(.dtpTo.Value, "yyyy-mm-dd") & "') AND ((SPITR.TRANSTYPE)='OUT'))"
                .lblRemarks.Caption = GrpCrit
                Call .ShowFilteredRecords
            End With
        Case "INVBCKSAL"
            With SPbckOrdSum
                .txtCriteria.Text = Qctr
                .lblRemarks.Caption = GrpCrit
                Call .ShowFilteredRecords
            End With
        Case "PRICELIST"
            With PRNmc
                .txtCriteria.Text = Qctr
                .lblRemarks.Caption = "Note: Filtered by " & GrpCrit
                Call .RefreshPriceList
            End With
        Case "SPINVENTORY"
            With SPInventory
                .txtSQL.Text = Qctr
                Call .ShowFilteredRecords
            End With
        Case "SPTOPSALES"
            With SPTopSales
                .txtSQL.Text = Left$(Qctr, Len(Qctr) - 1)
                .txtRemarks = "[FILTERED BY: " & GrpCrit & "] "
                Call .ShowFilteredRecords
            End With
        Case "DISBURSEMENT"
            With DisburseSumm
                .txtCriteria.Text = Left$(Qctr, Len(Qctr) - 1)
                .lblRemarks.Caption = "Filtered by " & GrpCrit
                Call .ShowFilteredRecords
            End With
        'New code her
        Case "arCheckSumm"
            With arCheckSumm
                .txtSQL.Text = Qctr
                Call .ShowFilteredRecords
            End With
    End Select
    Unload Me
 End Sub

Private Sub Form_Load()
    Dim Qry As String
    Dim X As Integer
    
    Me.Width = 6540
    Me.Height = 2520
        
    CurrentH = 2520
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier(0).Clear
            cboSupplier(1).Clear
            cboSupplier(2).Clear
            Do While Not .EOF
                cboSupplier(0).AddItem !SUPPLIERNAME
                cboSupplier(1).AddItem !SUPPLIERNAME
                cboSupplier(2).AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub

Private Sub Form_Resize()
    Me.Width = 6540
    Me.Height = CurrentH
End Sub


Private Sub txtFind_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strValid As String
    Dim y As String
    
    y = Trim(Str(cboLookIn(Index).ListIndex))
    strValid = "1234567890./-"
    
    If InStr(fldNdx, y) <> 0 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    Else
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
    
End Sub
