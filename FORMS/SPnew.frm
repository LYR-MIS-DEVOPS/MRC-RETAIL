VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPnew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Spare Part"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   Icon            =   "SPnew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin lvButton.lvButtons_H cmdWait 
      Height          =   615
      Left            =   2040
      TabIndex        =   24
      Top             =   1920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      Caption         =   "LOADING.............."
      CapAlign        =   2
      BackStyle       =   7
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
      Image           =   "SPnew.frx":08CA
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   4080
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   2400
      TabIndex        =   21
      Top             =   0
      Width           =   3735
      Begin VB.OptionButton OptSPType 
         Caption         =   "REPLACEMENT"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptSPType 
         Caption         =   "GENUINE"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSrch 
      Height          =   1695
      Left            =   2640
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2990
      _Version        =   393216
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
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
      Image           =   "SPnew.frx":0E64
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtGDesc 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtGPartNo 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "REPLACEMENT"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   5895
      Begin VB.TextBox txtRepl 
         Height          =   285
         Left            =   4680
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtSuper 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox txtDesc2 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   1035
         Width           =   3135
      End
      Begin VB.TextBox txtPartNo 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtDesc1 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   690
         Width           =   3135
      End
      Begin VB.TextBox txtCost 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   1365
         Width           =   3135
      End
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   1710
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SUPERSEDING"
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   19
         Top             =   2055
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DESCRIPTION 2"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   18
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PART NUMBER"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   15
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DESCRIPTION 1"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   705
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "COST"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   13
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "RETAIL PRICE"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   1725
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
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
      Image           =   "SPnew.frx":13FE
      cBack           =   -2147483633
   End
   Begin VB.Label lblGDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "DESCRIPTION"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblGPnum 
      Alignment       =   1  'Right Justify
      Caption         =   "GENUINE PART NUMBER"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   750
      Width           =   2175
   End
End
Attribute VB_Name = "SPnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim PrtSrch As Boolean, PrtBln As Boolean

Private Sub cmdCancel_Click()
    SPincoming.flxParts.TextMatrix(SPincoming.flxParts.Rows - 1, 1) = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If OptSPType(0).Value = True Then
        If txtPartNo.Text = "" Or txtDesc1.Text = "" Then
            MsgBox "Incomplete Entry!", vbExclamation + vbOKOnly, "Warning"
            Exit Sub
        End If
    Else
        If txtRepl.Text = "" Or txtGDesc.Text = "" Then
            MsgBox "Incomplete Entry!", vbExclamation + vbOKOnly, "Warning"
            Exit Sub
        End If
    End If
    With SPincoming
        .flxParts.TextMatrix(.flxParts.Rows - 1, 1) = txtPartNo.Text & IIf(OptSPType(1).Value = True, txtRepl.Text, "")
        .flxParts.TextMatrix(.flxParts.Rows - 1, 2) = txtDesc1.Text
        .flxParts.TextMatrix(.flxParts.Rows - 1, 6) = "New"
        .flxParts.TextMatrix(.flxParts.Rows - 1, 9) = .flxNewSP.Rows
        .flxNewSP.AddItem txtPartNo.Text & IIf(OptSPType(1).Value = True, txtRepl.Text, "") & vbTab & txtCost.Text & vbTab & txtPrice.Text & vbTab & txtDesc2.Text & vbTab & txtSuper.Text & vbTab & IIf(OptSPType(1).Value = True, "REPLACEMENT", "GENUINE")
    End With
    Unload Me
End Sub

Private Sub LoadSpareParts()
    Dim Qry As String
        
    If rsParts.State = adStateOpen Then Exit Sub
    Qry = "SELECT SOURCESP.* From SOURCESP ORDER BY SOURCESP.PARTNUMBER;"
    
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSrch.DataSource = rsParts
    With flxSrch
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = .Width
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
    End With
    cmdWait.Visible = False

End Sub


Private Sub Form_Load()
    Dim Qry As String, myPN As String
        
    OptSPType(0).Value = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub

Private Sub ClearEntries()
    txtGPartNo.Text = ""
    txtGDesc.Text = ""
    txtPartNo.Text = ""
    txtRepl.Text = ""
    txtDesc1.Text = ""
    txtDesc2.Text = ""
    txtCost.Text = ""
    txtPrice.Text = ""
    txtSuper.Text = ""
End Sub

Private Sub OptSPType_Click(Index As Integer)
    Dim myPN As String
    
    Call ClearEntries
    If OptSPType(0).Value = True Then
        lblGPnum.Visible = False
        txtGPartNo.Visible = False
        lblGDesc.Visible = False
        txtGDesc.Visible = False
        Frame1.Caption = "GENUINE PARTS"
        Frame1.Top = Frame2.Top + Frame2.Height + 100
        cmdOk.Top = Frame1.Top + Frame1.Height + 100
        cmdCancel.Top = cmdOk.Top
        txtRepl.Visible = False
        txtPartNo.Width = txtDesc1.Width
        Me.Height = cmdOk.Top + cmdOk.Height + 500
        myPN = SPincoming.flxParts.TextMatrix(SPincoming.flxParts.Rows - 1, 1)
        txtPartNo.Text = myPN
        txtPartNo.Alignment = 0
        txtPartNo.Locked = False
        txtDesc1.Locked = False
        txtSuper.Locked = False
    Else
        lblGPnum.Visible = True
        txtGPartNo.Visible = True
        lblGDesc.Visible = True
        txtGDesc.Visible = True
        Frame1.Caption = "REPLACEMENT PARTS"
        Frame1.Top = txtGDesc.Top + txtGDesc.Height + 100
        cmdOk.Top = Frame1.Top + Frame1.Height + 100
        cmdCancel.Top = cmdOk.Top
        txtRepl.Visible = True
        txtPartNo.Width = txtDesc1.Width - txtRepl.Width
        Me.Height = cmdOk.Top + cmdOk.Height + 500
        txtPartNo.Alignment = 1
        txtPartNo.Locked = True
        txtDesc1.Locked = True
        txtSuper.Locked = True
        myPN = SPincoming.flxParts.TextMatrix(SPincoming.flxParts.Rows - 1, 1)
        If InStr(myPN, "-R") > 0 Then
            txtGPartNo.Text = Left$(myPN, InStrRev(myPN, "-R") - 1)
            Call txtGPartNo_Change
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Call LoadSpareParts
    Timer1.Enabled = False
End Sub

Private Sub txtCost_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtPrice.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        txtDesc2.SetFocus
    End If
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtPrice.SetFocus
End Sub

Private Sub txtDesc1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtDesc2.SetFocus
    ElseIf KeyCode = vbKeyUp And OptSPType(1).Value = True Then
        txtRepl.SetFocus
    ElseIf KeyCode = vbKeyUp And OptSPType(0).Value = True Then
        txtPartNo.SetFocus
    End If
End Sub

Private Sub txtDesc1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDesc2.SetFocus
End Sub

Private Sub txtDesc2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtCost.SetFocus
    ElseIf KeyCode = vbKeyUp And OptSPType(1).Value = True Then
        txtRepl.SetFocus
    ElseIf KeyCode = vbKeyUp And OptSPType(0).Value = True Then
        txtDesc1.SetFocus
    End If
End Sub

Private Sub txtDesc2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtCost.SetFocus
End Sub

Private Sub txtGPartNo_Change()
    Dim i As Integer
    
    With flxSrch
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtGPartNo.Height + txtGPartNo.Top
        .Left = txtGPartNo.Left
        .Width = txtGPartNo.Width
        .Visible = True
        If txtGPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtGPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtGPartNo.Tag = Len(txtGPartNo.Text)
                PrtBln = True
            End If
            .TopRow = rsParts.AbsolutePosition - 1
            .Row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 5
            txtGPartNo.Text = rsParts!PARTNUMBER
            txtPartNo.Text = rsParts!PARTNUMBER
            'Insert Additional retrieval here
            txtGDesc.Text = .TextMatrix(.Row, 3)  'Description
            txtDesc1.Text = .TextMatrix(.Row, 3)  'Description
            txtSuper.Text = .TextMatrix(.Row, 4) 'Superseding
            txtGPartNo.SelStart = Val(txtGPartNo.Tag)
            txtGPartNo.SelLength = Len(txtGPartNo.Text)
            PrtBln = False
        Else
            txtGDesc.Text = ""  'Description
            txtDesc1.Text = ""  'Description
            txtSuper.Text = ""
            .Visible = False
        End If
    End With
End Sub

Private Sub txtGPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyDown
            txtGPartNo.Tag = "skiplostfocus"
            flxSrch.SetFocus
        Case vbKeyEscape
            flxSrch.Visible = False
            txtGPartNo.SetFocus
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
    End Select
End Sub

Private Sub txtGPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxSrch.Visible = False
        txtRepl.SetFocus
    End If
End Sub

Private Sub txtGPartNo_LostFocus()
    If txtGPartNo.Tag = "skiplostfocus" Then
        txtGPartNo.Tag = ""
        Exit Sub
    End If
    flxSrch.Visible = False
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtDesc1.SetFocus
    ElseIf KeyCode = vbKeyUp And OptSPType(0).Value = True Then
        txtGDesc.SetFocus
    End If
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDesc1.SetFocus
End Sub

Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown And OptSPType(1).Value = True Then
        cmdOk.SetFocus
    ElseIf KeyCode = vbKeyDown And OptSPType(0).Value = True Then
        txtSuper.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        txtCost.SetFocus
    End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If OptSPType(1).Value = True Then
            cmdOk.SetFocus
        Else
            txtSuper.SetFocus
        End If
    End If
End Sub

Private Sub txtRepl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtDesc2.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        txtGPartNo.SetFocus
    End If
End Sub

Private Sub txtRepl_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtDesc2.SetFocus
End Sub
