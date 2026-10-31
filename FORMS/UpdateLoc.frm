VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form UpdateLoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spare Parts Update Location"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   Icon            =   "UpdateLoc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSPID 
      Height          =   285
      Left            =   3480
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Update"
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
      Image           =   "UpdateLoc.frx":058A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtReason 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtNewLoc 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox txtPrevLoc 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtDate 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCtrlNo 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "DATE"
         Height          =   255
         Left            =   4680
         TabIndex        =   15
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "CTRL. NO."
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   255
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   3615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1365
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Close"
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
      Image           =   "UpdateLoc.frx":0B24
      cBack           =   -2147483633
   End
   Begin VB.Label Label7 
      Caption         =   "REASON"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "NEW LOCATION"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "DESCRIPTION"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "PREVIOUS LOCATION"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "PART NUMBER"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   855
      Width           =   1695
   End
End
Attribute VB_Name = "UpdateLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean
Dim Qry As String


Private Sub cmdClose_Click()
    Unload Me
End Sub

Function ValidateChangeLocation()
    Dim eMsg As String, Qry As String
    
    eMsg = ""
    
    If txtCtrlNo.Text = "" Then
        eMsg = "Incomplete Entry on field CTRL. NO."
    Else
       ' Qry = "SELECT TRANSFERLOC.TRANSFERNO From TRANSFERLOC WHERE (((TRANSFERLOC.TRANSFERNO)='" & txtCtrlNo.Text & "'));"
        'mySQL
        Qry = "SELECT TRANSFERLOC.TRANSFERNO From TRANSFERLOC WHERE TRANSFERLOC.TRANSFERNO='" & txtCtrlNo.Text & "';"
        
        If CtrlNumExist(Qry) Then
            eMsg = "CTRL. No. " & txtCtrlNo.Text & " already exist"
        End If
    End If
    If Not IsDate(txtDate.Text) Then eMsg = eMsg & vbCrLf & "Invalid Date Value"
    If txtNewLoc.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field New Location"
    If txtReason.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Reason"
    
    ValidateChangeLocation = eMsg
End Function

Private Sub cmdUpdate_Click()
    Dim rsTrans As New ADODB.Recordset
    Dim Ans As Integer
    Dim vMsg As String
    
    'Validate entry
    If txtSPID.Text = "" Then
        MsgBox "No Valid Part Number specified!", vbExclamation + vbOKOnly, "Validation Error"
        Exit Sub
    End If
'    If txtPrevLoc.Text = "" Then
'        Ans = MsgBox("Do you want to Save New Location?", vbQuestion + vbYesNo, "Save")
'        If Ans = vbNo Then Exit Sub
'        flxParts.TextMatrix(flxParts.Row, 4) = txtNewLoc.Text
'        'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & txtNewLoc.Text & "' " & _
'                       "WHERE (((SPAREPARTS.SPAREPARTID)=" & txtSPID.Text & "));"
'        'mySQL
'        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & txtNewLoc.Text & "' " & _
'                       "WHERE SPAREPARTS.SPAREPARTID=" & txtSPID.Text & ";"
'
'
'    Else
'        MsgBox "Only Spare Parts with No Location can be process here!", vbExclamation + vbOKOnly, "Location Exist"
'        Exit Sub
'    End If
'    MsgBox "New Location successfully updated!", vbInformation + vbOKOnly, txtPartNo.Text
    If txtPrevLoc.Text <> txtNewLoc.Text Then
        Ans = MsgBox("Do you want to Save New Location?", vbQuestion + vbYesNo, "Save")
        If Ans = vbNo Then Exit Sub
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & txtNewLoc.Text & "' " & _
                            "WHERE SPAREPARTS.SPAREPARTID=" & txtSPID.Text & ";"
    Else
        MsgBox "No changes for update", vbExclamation + vbOKOnly, "Location"
        Exit Sub
    End If
    
    
    txtPartNo.SetFocus
    txtPartNo.SelStart = 0
    txtPartNo.SelLength = Len(txtPartNo.Text)
End Sub

Private Sub flxParts_Click()
    FromFlx = True
End Sub

Private Sub flxParts_EnterCell()
    With flxParts
        FromFlx = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxParts.Visible = False
        txtNewLoc.SetFocus
    End If
End Sub

Private Sub flxParts_LostFocus()
    flxParts.Visible = False
End Sub


Private Sub Form_Load()
    MotoRep.sBar.Panels(6).Text = Me.Caption
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"

    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = txtPartNo.Width
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = txtPartNo.Width
        .ColWidth(3) = 0
        .ColWidth(4) = 0
    End With
    txtDate.Text = Date
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub


Private Sub txtNewLoc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdUpdate.SetFocus
End Sub

Private Sub txtPrevLoc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdUpdate.SetFocus
End Sub

Private Sub txtPartNo_Change()
    Dim mRow As Integer
    
    With flxParts
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            If Not FromFlx Then     'top row only if user type @ txtpartno
                .TopRow = rsParts.AbsolutePosition - 1
                FromFlx = False
            End If
            .Row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 3
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            txtDesc.Text = flxParts.TextMatrix(flxParts.Row, 3) 'Description
            txtPrevLoc.Text = flxParts.TextMatrix(flxParts.Row, 4)  'Location
            txtSPID.Text = flxParts.TextMatrix(flxParts.Row, 1) 'SP ID
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            txtNewLoc.Text = ""
            txtReason.Text = ""
            PrtBln = False
        Else
            txtDesc.Text = ""
            txtPrevLoc.Text = ""
            txtSPID.Text = ""
            .Visible = False
        End If
    End With
    
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdExit.SetFocus
            flxParts.Visible = False
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
        Case vbKeyDown
            flxParts.Visible = True
            flxParts.SetFocus
        Case Else
            FromFlx = False
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxParts.Visible = False
        txtNewLoc.SetFocus
    End If
End Sub
