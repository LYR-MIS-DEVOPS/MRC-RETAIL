VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmClubInfo 
   Caption         =   "CLUB INFORMATION"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   7680
      TabIndex        =   21
      Text            =   "HoldBackColor"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "CLUB INFORMATION"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      Begin VB.TextBox txtClubSEC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   7
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtStreet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   6
         Top             =   2160
         Width           =   5295
      End
      Begin VB.ComboBox cboBarangay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   5
         Top             =   1800
         Width           =   5295
      End
      Begin VB.ComboBox cboCITY 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   4
         Top             =   1440
         Width           =   5295
      End
      Begin VB.ComboBox cboProvince 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   5295
      End
      Begin VB.ComboBox cboRegion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox txtClubName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label18 
         Caption         =   "CLUB SEC. NO.:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2625
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "ADDRESS STREET:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2265
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "BARANGAY:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1875
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "CITY/TOWN:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1515
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "REGION:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   800
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "PROVINCE:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "CLUB NAME:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   400
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Add New"
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
      Image           =   "frmClubInfo.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1620
      TabIndex        =   16
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "frmClubInfo.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4740
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Save"
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
      Image           =   "frmClubInfo.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmClubInfo.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Edit"
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
      Image           =   "frmClubInfo.frx":1668
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmClubInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Qry As String
Dim Ans As String
Dim mainRs As New ADODB.Recordset
Dim i As Long, j As Long, myCCID As Long
Dim StatSrch As Boolean
Dim strPartial As String, strTotal As String
Dim PrtBln As Boolean, PrtSrch As Boolean, ProperExit As Boolean
Public MySource As String, myClubType As String
Dim formAction As String

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    txtClubName.Locked = DataEntry
    cboRegion.Locked = DataEntry
    cboProvince.Locked = DataEntry
    cboCITY.Locked = DataEntry
    cboBarangay.Locked = DataEntry
    txtStreet.Locked = DataEntry
    txtClubSEC.Locked = DataEntry
    
    If txtClubName.Locked Then
        txtClubName.BackColor = txtHoldGrayBackColor.BackColor
        cboRegion.BackColor = txtHoldGrayBackColor.BackColor
        cboProvince.BackColor = txtHoldGrayBackColor.BackColor
        cboCITY.BackColor = txtHoldGrayBackColor.BackColor
        cboBarangay.BackColor = txtHoldGrayBackColor.BackColor
        txtStreet.BackColor = txtHoldGrayBackColor.BackColor
        txtClubSEC.BackColor = txtHoldGrayBackColor.BackColor

    Else
        txtClubName.BackColor = txtHoldWhiteBackColor.BackColor
        cboRegion.BackColor = txtHoldWhiteBackColor.BackColor
        cboProvince.BackColor = txtHoldWhiteBackColor.BackColor
        cboCITY.BackColor = txtHoldWhiteBackColor.BackColor
        cboBarangay.BackColor = txtHoldWhiteBackColor.BackColor
        txtStreet.BackColor = txtHoldWhiteBackColor.BackColor
        txtClubSEC.BackColor = txtHoldWhiteBackColor.BackColor
    End If
    
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->FIND
'2 ->EXIT
'3 ->SAVE
'4 ->CANCEL

    'hide all button
    For i = 0 To 4
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(3).Visible = True
        cmdTrans(4).Visible = True
    Else
        For i = 0 To 4
            cmdTrans(i).Visible = True
            cmdTrans(3).Visible = False
            cmdTrans(4).Visible = False
        Next i
    End If
End Sub

Public Sub ShowRecord(MyID As Long)

'    Call ConnectOnline
'    If modConnection.TestConnection() = False Then Exit Sub
    myCCID = 0
    myCCID = MyID
    Qry = "call sp_promotionalonline.sp_promocard_getClubDetails(" & Val(MyID) & ")"
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtClubName.Text = IIf(IsNull(!CLUBNAME), "", !CLUBNAME)
            cboRegion.Text = IIf(IsNull(!REGION), "", !REGION)
            cboProvince.Text = IIf(IsNull(!province), "", !province)
            cboCITY.Text = IIf(IsNull(!CITYTOWN), "", !CITYTOWN)
            cboBarangay.Text = IIf(IsNull(!barangay), "", !barangay)
            txtStreet.Text = IIf(IsNull(!STREET), "", !STREET)
            txtClubSEC.Text = IIf(IsNull(!CLUBSECNO), "", !CLUBSECNO)
        End If
        .Close
    End With
End Sub
Function ValidateEntry() As Boolean
  ValidateEntry = False
    
    If Trim(txtClubName.Text) = "" Then
        MsgBox "CLUB NAME IS REQUIRED!", vbExclamation, "Validation"
        txtClubName.SetFocus
        Exit Function
    End If
    
    If Trim(cboRegion.Text) = "" Then
        MsgBox "REGION IS REQUIRED!", vbExclamation, "Validation"
        cboRegion.SetFocus
        Exit Function
    End If
        
    If Trim(cboProvince.Text) = "" Then
        MsgBox "PROVINCE IS REQUIRED!", vbExclamation, "Validation"
        cboProvince.SetFocus
        Exit Function
    End If
    
    If Trim(cboCITY.Text) = "" Then
        MsgBox "CITY/TOWN IS REQUIRED!", vbExclamation, "Validation"
        cboCITY.SetFocus
        Exit Function
    End If
    
    If Trim(cboBarangay.Text) = "" Then
        MsgBox "BARANGAY IS REQUIRED!", vbExclamation, "Validation"
        cboBarangay.SetFocus
        Exit Function
    End If
    
    If Trim(txtStreet.Text) = "" Then
        MsgBox "STREET IS REQUIRED!", vbExclamation, "Validation"
        txtStreet.SetFocus
        Exit Function
    End If
    
    ValidateEntry = True

End Function

Sub clearform()

    txtClubName.Text = ""
    cboRegion.Text = ""
    cboProvince.Text = ""
    cboCITY.Text = ""
    cboBarangay.Text = ""
    txtStreet.Text = ""
    txtClubSEC.Text = ""
    
End Sub

Public Sub LoadRegion()

    Qry = "call sp_mrcretails.sp_geomas_phil_list('','')"
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If formAction = "add" Then
                cboRegion.Clear
            End If
            Do While Not .EOF
                cboRegion.AddItem !Name
                .MoveNext
            Loop
        End If
        .Close
    End With
End Sub

Private Sub cboBarangay_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboBarangay
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBarangay-Change()")
End Sub

Private Sub cboBarangay_Click()
If cboBarangay.Text <> "" Then
    txtStreet.Text = ""
    txtStreet.SetFocus
End If
End Sub

Private Sub cboBarangay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboBarangay_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboBarangay_Click
    End If
End Sub

Private Sub cboCITY_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboCITY
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCITY-Change()")
End Sub

Private Sub cboCITY_Click()
    Dim barangay As String
    Qry = "SELECT code_1to6 FROM sp_mrcretails.geomas_phil Where NAME = '" & Trim(cboCITY.Text) & "' "
    barangay = GetFieldValue(Qry)
    
    Qry = "call sp_mrcretails.sp_geomas_phil_list('" & "barangay" & "','" & barangay & "')"
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboBarangay.Clear
            txtStreet.Text = ""
            Do While Not .EOF
                cboBarangay.AddItem !Name
                .MoveNext
            Loop
        End If
        .Close
    End With
    cboBarangay.SetFocus
End Sub

Private Sub cboCITY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboCITY_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboCITY_Click
    End If
End Sub

Private Sub cboProvince_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboProvince
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboProvince-Change()")
End Sub

Private Sub cboProvince_Click()
    Dim city As String
    Qry = "SELECT code_1to4 FROM sp_mrcretails.geomas_phil Where NAME = '" & Trim(cboProvince.Text) & "' "
    city = GetFieldValue(Qry)
    
    Qry = "call sp_mrcretails.sp_geomas_phil_list('" & "city" & "','" & city & "')"
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboCITY.Clear
            cboBarangay.Clear
            txtStreet.Text = ""
            
            Do While Not .EOF
                cboCITY.AddItem !Name
                .MoveNext
            Loop
        End If
        .Close
    End With
    cboCITY.SetFocus
End Sub

Private Sub cboProvince_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboProvince_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboProvince_Click
    End If
End Sub

Private Sub cboRegion_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboRegion
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboRegion-Change()")
End Sub

Private Sub cboRegion_Click()
    Dim province As String
    Qry = "SELECT code_1to2 FROM sp_mrcretails.geomas_phil Where NAME = '" & Trim(cboRegion.Text) & "' "
    province = GetFieldValue(Qry)
    
    Qry = "call sp_mrcretails.sp_geomas_phil_list('" & "province" & "','" & province & "')"
    With mainRs
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            
            cboProvince.Clear
            cboCITY.Clear
            cboBarangay.Clear
            txtStreet.Text = ""
            
            Do While Not .EOF
                cboProvince.AddItem !Name
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    cboProvince.SetFocus
End Sub

Private Sub cboRegion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cboRegion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cboRegion_Click
    End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdTrans_Click(Index As Integer)
    If Index = 0 Then
        formAction = "add"
        Call clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        Call LoadRegion
    ElseIf Index = 1 Then
        formAction = "edit"
        Call LockData(True)
        Call SetButtonVisibility(True)
        Call LoadRegion
    ElseIf Index = 2 Then
        formAction = "EXIT"
        ProperExit = True
        formAction = ""
        Unload Me
    ElseIf Index = 3 Then
        If ValidateEntry() = False Then Exit Sub
    
        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        
        If Ans = vbNo Then Exit Sub
        
'        Call ConnectOnline
'        If modConnection.TestConnection() = False Then Exit Sub
        
        If formAction = "add" Then
            Qry = "call sp_promotionalonline.sp_promocard_clubinfo_create('" & Trim(txtClubName.Text) & "','" & Trim(cboRegion.Text) & "','" & Trim(cboProvince.Text) & "','" & Trim(cboCITY.Text) & "','" & Trim(cboBarangay.Text) & "','" & Trim(txtStreet.Text) & "','" & Trim(txtClubSEC.Text) & "','" & Trim(myClubType) & "'," & Val(UserID) & ")"
            With mainRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    If !level = "Granted" Then
                        frmPromoCard.txtCCID.Text = IIf(IsNull(!CCID), 0, !CCID)
                        frmPromoCard.cboCustClubName.Text = txtClubName.Text
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "SAVE"
                    Else
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "SAVE"
                        .Close
                        Exit Sub
                    End If
                End If
                .Close
            End With
        ElseIf formAction = "edit" Then
            Qry = "call sp_promotionalonline.sp_promocard_clubinfo_update(" & Val(myCCID) & ",'" & Trim(txtClubName.Text) & "','" & Trim(cboRegion.Text) & "','" & Trim(cboProvince.Text) & "','" & Trim(cboCITY.Text) & "','" & Trim(cboBarangay.Text) & "','" & Trim(txtStreet.Text) & "','" & Trim(txtClubSEC.Text) & "'," & Val(UserID) & ")"
            With mainRs
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    If !level = "Granted" Then
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "SAVE"
                    Else
                        MsgBox "" & !message & "", vbInformation + vbOKOnly, "SAVE"
                        .Close
                        Exit Sub
                    End If
                End If
                .Close
            End With
        End If
        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
    ElseIf Index = 4 Then
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        If Ans = vbNo Then Exit Sub
        Call SetButtonVisibility(False)
        Call LockData(False)
        If formAction = "add" Then
            cmdTrans(0).Visible = True
            cmdTrans(1).Visible = False
        ElseIf formAction = "edit" Then
            cmdTrans(0).Visible = False
            cmdTrans(1).Visible = True
        End If
        formAction = ""
    End If
End Sub

Private Sub Form_Load()
    Call LockData(False)
        '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(3).Left = cmdTrans(0).Left
        cmdTrans(3).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(4).Left = cmdTrans(2).Left
        cmdTrans(4).Top = cmdTrans(2).Top
    End If
End Sub

Private Sub txtClubName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub txtClubSEC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub txtStreet_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub
