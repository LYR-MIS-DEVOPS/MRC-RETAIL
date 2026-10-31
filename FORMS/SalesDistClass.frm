VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SalesDistClass 
   Caption         =   "Update Classification"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   4680
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Tag             =   "xyf"
      Top             =   3960
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Tag             =   "xyf"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Tag             =   "xyf"
      Top             =   4110
      Width           =   2775
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update All Parts Classification to:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   4080
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid FlxParts 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   480
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      FormatString    =   "      |<PART NUMBER          |<DESCRIPTION                               |<CLASSIFICATION             "
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   90
      Width           =   1935
   End
   Begin MSComCtl2.UpDown UpDRef 
      Height          =   315
      Left            =   1845
      TabIndex        =   6
      Tag             =   "xyf"
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      Max             =   2
      Enabled         =   -1  'True
   End
   Begin VB.Label lblRefSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIND PARTS DESC."
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Tag             =   "xyf"
      Top             =   180
      Width           =   1605
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   9060
      m_MDI_Height    =   5580
   End
   Begin VB.Menu mnuUpdate 
      Caption         =   "Update Classification"
      Visible         =   0   'False
      Begin VB.Menu mnuCto 
         Caption         =   "Update Classification to"
         Begin VB.Menu mnuClsHonda 
            Caption         =   "HONDA GENUINE"
         End
         Begin VB.Menu mnuClsYamaha 
            Caption         =   "YAMAHA GENUINE"
         End
         Begin VB.Menu mnuClsSuzuki 
            Caption         =   "SUZUKI GENUINE"
         End
         Begin VB.Menu mnuClsReplacement 
            Caption         =   "REPLACEMENT"
         End
         Begin VB.Menu mnuClsAccessories 
            Caption         =   "ACCESSORIES"
         End
         Begin VB.Menu mnuClsPerformance 
            Caption         =   "PERFORMANCE"
         End
         Begin VB.Menu mnuClsApparel 
            Caption         =   "RIDER APPARELS"
         End
         Begin VB.Menu mnuClsTires 
            Caption         =   "TIRES"
         End
         Begin VB.Menu mnuClsOil 
            Caption         =   "OIL"
         End
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Item"
      End
   End
End
Attribute VB_Name = "SalesDistClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public spQry As String
Dim ActRow As Integer

Private Sub cboClass_GotFocus()
    'Call PopupMenu(mnuDC)
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub cmdClose_Click()
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
                myW1 = "SPAREPARTS.PARTNUMBER Like '%" & MySP(i) & "%'"
            ElseIf UpDRef.Value = 1 Then
                myW1 = "SPAREPARTS.DESCRIPTION Like '%" & MySP(i) & "%'"
            ElseIf UpDRef.Value = 2 Then
                myW1 = "supplier.SUPPLIERNAME LIKE '%" & MySP(i) & "%'"
            End If
            
        Else
            'myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            If UpDRef.Value = 0 Then
                myW1 = myW1 & " AND SPAREPARTS.PARTNUMBER Like '%" & MySP(i) & "%'"
            ElseIf UpDRef.Value = 1 Then
                myW1 = myW1 & " AND SPAREPARTS.DESCRIPTION Like '%" & MySP(i) & "%'"
            ElseIf UpDRef.Value = 2 Then
                myW1 = "supplier.SUPPLIERNAME LIKE '%" & MySP(i) & "%'"
            End If
        End If
    Next i
'    myW = myW1
'    If ShowOpt Then
'        If txtModel.Text <> "" Then
'            Mdl = Split(txtModel.Text, " ")
'            myW2 = ""
'            For i = LBound(Mdl) To UBound(Mdl)
'                If myW2 = "" Then
'                    myW2 = "(SPAREPARTS.APPLICMODEL) Like '%" & Mdl(i) & "%'"
'                Else
'                    myW2 = myW2 & " AND (SPAREPARTS.APPLICMODEL) Like '%" & Mdl(i) & "%'"
'                End If
'            Next i
'            myW = "(" & myW & ") AND (" & myW2 & ")"
'        End If
'        If cboSPType.ListIndex >= 1 Then
'            myW = "(" & myW & ") AND ((SPAREPARTS.REMARKS)='" & cboSPType.Text & "')"
'        End If
'    End If
    '***Access
    'MySel = "SELECT SPAREPARTS.QTYONHAND, SPAREPARTS.PARTNUMBER, [DESCRIPTION] & IIf(IsNull([DESCRIPTION2]),'',' ' & [DESCRIPTION2]) AS MYDESC, SPAREPARTS.APPLICMODEL, SPAREPARTS.STOREPRICE, SPAREPARTS.MAXDISCOUNT, SPAREPARTS.LOCATION, SPAREPARTS.SPAREPARTID, SPAREPARTS.COST From SPAREPARTS WHERE (" & myW & ");"
    
    'mySQL
    If UpDRef.Value <= 1 Then
        MySel = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.REMARKS, SPAREPARTS.SPAREPARTID From SPAREPARTS WHERE " & myW1 & ";"
    Else
        MySel = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.REMARKS, SPAREPARTS.SPAREPARTID " & _
                "FROM spareparts INNER JOIN supplier ON (spareparts.SUPPLIERCODE = supplier.SUPPLIERCODE)"
    End If

    With rsSP
        .Open MySel, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxParts.Rows = 1
            Do While Not .EOF
                flxParts.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !REMARKS & vbTab & !SPAREPARTID
                .MoveNext
            Loop
        Else
            flxParts.Clear
            flxParts.Rows = 1
            flxParts.Rows = 2
        End If
        .Close
    End With
    Call AddInCombo
    If flxParts.Rows >= 2 Then
        If flxParts.TextMatrix(1, 3) <> "" Then
            flxParts.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
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



Private Sub cmdUpdate_Click()
    Dim Ans As Integer, i As Integer
    
    If cboClass.ListIndex = -1 Then
        MsgBox "Please select a Valid SpareParts Classification!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    Ans = MsgBox("Do you want update All SpareParts classification to " & cboClass.Text & "?", vbQuestion + vbYesNo, cboClass.Text)
    If Ans = vbNo Then Exit Sub
    
    For i = 1 To flxParts.Rows - 1
        If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
            LYRcon.Execute "UPDATE spareparts SET REMARKS = '" & cboClass.Text & "' Where spareparts.SPAREPARTID = " & Val(flxParts.TextMatrix(i, 4)) & ";"
        ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
            BrchConMU.Execute "UPDATE spareparts SET REMARKS = '" & cboClass.Text & "' Where spareparts.SPAREPARTID = " & Val(flxParts.TextMatrix(i, 4)) & ";"
        ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
            BrchConMD.Execute "UPDATE spareparts SET REMARKS = '" & cboClass.Text & "' Where spareparts.SPAREPARTID = " & Val(flxParts.TextMatrix(i, 4)) & ";"
        ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
            BrchConYM.Execute "UPDATE spareparts SET REMARKS = '" & cboClass.Text & "' Where spareparts.SPAREPARTID = " & Val(flxParts.TextMatrix(i, 4)) & ";"
        ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
            BrchConVAN.Execute "UPDATE spareparts SET REMARKS = '" & cboClass.Text & "' Where spareparts.SPAREPARTID = " & Val(flxParts.TextMatrix(i, 4)) & ";"
        ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
            BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = '" & cboClass.Text & "' Where spareparts.SPAREPARTID = " & Val(flxParts.TextMatrix(i, 4)) & ";"
        End If
        flxParts.TextMatrix(i, 3) = cboClass.Text
    Next i

    MsgBox "Spare Parts successfully updated to " & cboClass.Text
    
End Sub

Private Sub flxParts_EnterCell()
    ActRow = flxParts.Row
End Sub

Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If Val(flxParts.TextMatrix(flxParts.Row, 4)) > 0 Then
            Call PopupMenu(mnuUpdate)
        End If
    End If

End Sub



Private Sub Form_Load()
        
    'cboPartDesc.AddItem "DESCRIPTION"
    'cboPartDesc.AddItem "PART NUMBER"
    'cboPartDesc.ListIndex = 0
    With cboClass
        .AddItem "HONDA GENUINE"
        .AddItem "YAMAHA GENUINE"
        .AddItem "SUZUKI GENUINE"
        .AddItem "REPLACEMENT"
        .AddItem "ACCESSORIES"
        .AddItem "PERFORMANCE"
        .AddItem "RIDER APPARELS"
        .AddItem "TIRES"
        .AddItem "OIL"
    End With
    
    If spQry <> "" Then Timer1.Enabled = True
End Sub

Private Sub mnuclsAccessories_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to ACCESSORIES?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'ACCESSORIES' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "ACCESSORIES"
    MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully updated to ACCESSORIES"

End Sub

Private Sub mnuClsApparel_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to RIDER APPARELS?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'RIDER APPARELS' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "RIDER APPARELS"
    MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully updated to RIDER APPARELS"


End Sub


Private Sub mnuClsHonda_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to HONDA GENUINE?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'HONDA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "HONDA GENUINE"
    MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully updated to HONDA GENUINE"

End Sub

Private Sub mnuClsOil_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to OIL?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'OIL' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "OIL"
    MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully updated to OIL"

End Sub

Private Sub mnuClsPerformance_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to PERFORMANCE?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'PERFORMANCE' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "PERFORMANCE"
    MsgBox flxParts.TextMatrix(ActRow, 2) & " successfully updated to PERFORMANCE"

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SalesDistribution.SDCloaded = False
End Sub

Private Sub mnuClsTires_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to TIRES?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'TIRES' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "TIRES"
    MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully updated to TIRES"

End Sub

Private Sub mnuRemove_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to remove part number " & flxParts.TextMatrix(ActRow, 1) & "?", vbQuestion + vbYesNo, "Remove")
    If Ans = vbNo Then Exit Sub
    
    If flxParts.Rows = 2 And ActRow = 1 Then
        flxParts.Rows = 1
        flxParts.Rows = 2
    Else
        flxParts.RemoveItem ActRow
    End If
    
End Sub

Private Sub mnuClsReplacement_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to REPLACEMENT?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'REPLACEMENT' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "REPLACEMENT"
    MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully updated to REPLACEMENT"

End Sub

Private Sub mnuClsSuzuki_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to SUZUKI GENUINE?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'SUZUKI GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "SUZUKI GENUINE"
    MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully updated to SUZUKI GENUINE"

End Sub

Private Sub mnuClsYamaha_Click()
    Dim Ans As Integer, i As Integer, MyID As Long
    
    Ans = MsgBox("Do you want to update " & flxParts.TextMatrix(ActRow, 1) & " classification to YAMAHA GENUINE?", vbQuestion + vbYesNo, "Update")
    If Ans = vbNo Then Exit Sub
    
    MyID = Val(flxParts.TextMatrix(ActRow, 4))
    'update DB
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        LYRcon.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        BrchConMU.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        BrchConMD.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        BrchConYM.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        BrchConVAN.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        BrnConMRC2.Execute "UPDATE spareparts SET REMARKS = 'YAMAHA GENUINE' Where spareparts.SPAREPARTID = " & MyID & ";"
    Else
        MsgBox "You have not selected any Branch!" & vbCrLf & "Updating Failed!", vbExclamation + vbOKOnly, "No Branch"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 3) = "YAMAHA GENUINE"
    MsgBox flxParts.TextMatrix(ActRow, 1) & " successfully updated to YAMAHA GENUINE"

End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxParts
        .ColWidth(0) = 300
        .ColWidth(1) = .Width * 0.24
        .ColWidth(2) = .Width * 0.46
        .ColWidth(3) = .Width * 0.23
        .ColWidth(4) = 0
    End With
        
End Sub

Private Sub Timer1_Timer()
    Dim rsParts As New ADODB.Recordset
    
    
    If SalesDistribution.cboBranch.ListIndex = 0 Then     'MOTO-R MAIN"
        rsParts.Open spQry, LYRcon, adOpenKeyset, adLockOptimistic
    ElseIf SalesDistribution.cboBranch.ListIndex = 1 Then 'MOTO-R ULAS"
        rsParts.Open spQry, BrchConMU, adOpenKeyset, adLockOptimistic
    ElseIf SalesDistribution.cboBranch.ListIndex = 2 Then 'MC PARTS DEPOT"
        rsParts.Open spQry, BrchConMD, adOpenKeyset, adLockOptimistic
    ElseIf SalesDistribution.cboBranch.ListIndex = 3 Then 'YAMAHA MOTORHUB"
        rsParts.Open spQry, BrchConYM, adOpenKeyset, adLockOptimistic
    ElseIf SalesDistribution.cboBranch.ListIndex = 4 Then  'VAN SELLING"
        rsParts.Open spQry, BrchConVAN, adOpenKeyset, adLockOptimistic
    ElseIf SalesDistribution.cboBranch.ListIndex = 5 Then 'MRC2 GANTE"
        rsParts.Open spQry, BrnConMRC2, adOpenKeyset, adLockOptimistic
    End If
    With rsParts
        pBar.Value = 0
        pBar.Max = .RecordCount
        pBar.Visible = True
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            Do While Not .EOF
                flxParts.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !REMARKS & vbTab & !SPAREPARTID
                .MoveNext
                pBar = pBar + 1
                DoEvents
            Loop
        End If
        .Close
        pBar.Visible = False
    End With
    
    Timer1.Enabled = False
End Sub

Private Sub UpDRef_Change()
    If UpDRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    ElseIf UpDRef.Value = 1 Then
        lblRefSearch.Caption = "FIND PARTS DESC."
    ElseIf UpDRef.Value = 2 Then
        lblRefSearch.Caption = "FIND SUPPLIER"
    End If

End Sub
