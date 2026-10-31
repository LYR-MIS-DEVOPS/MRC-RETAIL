VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form EMPbranch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "On-Account Customer"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   6510
   Icon            =   "EMPbranch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   4800
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   4320
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   4800
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   4320
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6045
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   5400
         TabIndex        =   29
         Top             =   2060
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Increment       =   15
         Max             =   180
      End
      Begin VB.TextBox txtCLimit 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2395
         Width           =   3150
      End
      Begin VB.OptionButton OptType 
         Caption         =   "INDIVIDUAL"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   28
         Top             =   300
         Width           =   1815
      End
      Begin VB.OptionButton OptType 
         Caption         =   "COMPANY"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   27
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtOSBal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   7
         Top             =   2730
         Width           =   3150
      End
      Begin VB.TextBox txtCustID 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1055
         Width           =   3150
      End
      Begin VB.TextBox txtCustomer 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   3150
      End
      Begin VB.TextBox txtTerms 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   5
         Top             =   2060
         Width           =   2895
      End
      Begin VB.TextBox txtConPers 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1725
         Width           =   3150
      End
      Begin VB.TextBox txtContactNo 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1390
         Width           =   3150
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER TYPE:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   26
         Top             =   330
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADDRESS"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1085
         Width           =   2025
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OUTSTANDING BALANCE"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   17
         Top             =   2790
         Width           =   1965
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CREDIT LIMIT"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   8
         Left            =   840
         TabIndex        =   12
         Top             =   2455
         Width           =   1545
      End
      Begin VB.Label lblCustName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label lblContactPerson 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MANAGER"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1785
         Width           =   2025
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NOS."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   9
         Top             =   1450
         Width           =   2025
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TERMS"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   7
         Left            =   615
         TabIndex        =   8
         Top             =   2120
         Width           =   1770
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPbranch.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPbranch.frx":1BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPbranch.frx":3CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPbranch.frx":4BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPbranch.frx":5A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPbranch.frx":696C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "Add &New"
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
      Image           =   "EMPbranch.frx":99EE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   21
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "EMPbranch.frx":9F88
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   22
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "E&xit Form"
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
      Image           =   "EMPbranch.frx":A522
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "EMPbranch.frx":AABC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7200
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "EMPbranch.frx":B056
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   25
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Find"
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
      Image           =   "EMPbranch.frx":B5F0
      cBack           =   -2147483633
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ON-ACCOUNT"
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
      Left            =   3315
      TabIndex        =   15
      Top             =   540
      Width           =   3000
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
      Left            =   240
      TabIndex        =   13
      Top             =   545
      Width           =   3500
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
      Left            =   270
      TabIndex        =   14
      Top             =   570
      Width           =   3500
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ON-ACCOUNT"
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
      Left            =   3345
      TabIndex        =   16
      Top             =   570
      Width           =   3000
   End
   Begin VB.Menu mnuBranch 
      Caption         =   "Branch"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "EMPbranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBranch As New ADODB.Recordset
Dim AddNew As Boolean
Dim BtnIndx As Integer

Sub ClearEntries()
    
    txtCustID.Text = ""
    txtCustomer.Text = ""
    txtAddress.Text = ""
    txtContactNo.Text = ""
    txtConPers.Text = ""
    txtTerms.Text = ""
    txtCLimit.Text = ""
    txtOSBal.Text = ""

End Sub

Sub LockData(Bln As Boolean)
    txtCustID.Locked = Bln
    txtCustomer.Locked = Bln
    txtAddress.Locked = Bln
    txtContactNo.Locked = Bln
    txtConPers.Locked = Bln
    txtTerms.Locked = Bln
    txtCLimit.Locked = Bln
    txtOSBal.Locked = Bln
End Sub

Sub SaveBranch()
    Dim x As Integer
    Dim Qry As String
    
    If AddNew Then
        rsBranch.Open "BRANCH", LYRcon, adOpenKeyset, adLockOptimistic
        rsBranch.AddNew
    Else
        Qry = "SELECT BRANCH.* From BRANCH " & _
              "WHERE (((BRANCH.BRANCHID)=" & txtEntry(0).Text & "));"
        rsBranch.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        rsBranch.MoveFirst
    End If
    With rsBranch
        For x = 1 To 7
            .Fields(x) = txtEntry(x).Text
        Next x
        .Fields("UpdatedBy") = UserID
        .Fields("LastUpdate") = Date
        .Update
        If AddNew Then txtEntry(0).Text = !BranchID
        .Close
    End With
End Sub

Sub ShowRecord(BranchID As Integer)
    Dim Qry As String
    Dim x As Integer
    
    Qry = "SELECT BRANCH.* From BRANCH " & _
          "WHERE (((BRANCH.BRANCHid)=" & BranchID & "));"
    rsBranch.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    With rsBranch
        .MoveFirst
        For x = 0 To 7
            txtEntry(x).Text = IIf(IsNull(.Fields(x)), "", .Fields(x))
        Next x
        .Close
    End With
End Sub

Function ValidateEntry() As String
    Dim myMsg As String
    
    myMsg = ""
    If txtEntry(1).Text = "" Then myMsg = "No data on field BRANCH NAME"
    If txtEntry(2).Text = "" Then myMsg = myMsg & vbCrLf & "No data on field MANAGER"
    If txtEntry(3).Text = "" Then myMsg = myMsg & vbCrLf & "No data on field ADDRESS"
    
    ValidateEntry = myMsg
End Function

Private Sub BtnSaveCancel(Bln As Boolean)
    Dim i As Integer
    
    If Bln Then
        For i = 0 To 3
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
        cmdTrans(4).Left = 0
        cmdTrans(5).Left = cmdTrans(4).Width
    Else
        For i = 0 To 3
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim vMsg As String, QrySrch As String
    Dim BranchID As Integer, Ans As Integer
    
    Select Case Index
        Case 0  '   "new"
            If DispBtn Then
                Call BtnSaveCancel(True)
            Else
                BtnIndx = 3
                tmrFloat.Enabled = True
            End If
                
            Call LockData(False)
            Call ClearEntries
            AddNew = True
            txtEntry(1).SetFocus
        Case 1      '"edit"
            If txtEntry(0).Text = "" Then
                MsgBox "No records found!", vbExclamation + vbOKOnly, "Warning"
                txtEntry(1).SetFocus
                Exit Sub
            End If
            If DispBtn Then
                Call BtnSaveCancel(True)
            Else
                BtnIndx = 3
                tmrFloat.Enabled = True
            End If
            Call LockData(False)
            AddNew = False
            txtEntry(1).SetFocus
        Case 2          '"find"
            QrySrch = "SELECT BRANCH.BRANCHID, BRANCH.BRANCHNAME " & _
                      "From BRANCH ORDER BY BRANCH.BRANCHNAME;"

            frmSearch.QrySearch = QrySrch
            frmSearch.lblCap.Caption = "Search for Branch"
            Set frmSearch.FrmSource = EMPbranch
            frmSearch.Show vbModal
        Case 3          '"quit"
            MotoRep.tbMain.Visible = True
            Unload Me
        
        Case 4              '"save"
            vMsg = ValidateEntry
            If vMsg <> "" Then
                MsgBox vMsg, vbOKOnly + vbCritical, "Validation Error"
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Call SaveBranch
            MsgBox "Record successfully saved!", vbInformation + vbOKOnly, "Save"
            Call LockData(True)
            txtEntry(1).SetFocus
            If DispBtn Then
                Call BtnSaveCancel(False)
            Else
                BtnIndx = 5
                tmrFloatRev.Enabled = True
            End If
        Case 5          '"cancel"
            Ans = MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                Call ClearEntries
            Else
                BranchID = txtEntry(0).Text
                Call ClearEntries
                Call ShowRecord(BranchID)
            End If
            Call LockData(True)
            txtEntry(1).SetFocus
            If DispBtn Then
                Call BtnSaveCancel(False)
            Else
                BtnIndx = 5
                tmrFloatRev.Enabled = True
            End If
    End Select

End Sub

Private Sub Form_Load()
    
    Call LockData(True)
    'Call tbSaveCancel(False)
    
    MotoRep.sBar.Panels(6).Text = "On-Account Customer Entry Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub OptType_Click(Index As Integer)
    If Index = 0 Then
        lblCustName.Caption = "COMPANY NAME"
        lblContactPerson.Caption = "MANAGER"
    Else
        lblCustName.Caption = "CUSTOMER NAME"
        lblContactPerson.Caption = "EMPLOYER"
    End If
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then 'last index + 1
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 4 Then
            tmrDropRev.Enabled = False
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 4
            cmdTrans(4).Top = -375
            cmdTrans(4).Left = 0
            cmdTrans(4).Visible = True
            cmdTrans(5).Top = -375
            cmdTrans(5).Left = cmdTrans(4).Width
            cmdTrans(5).Visible = True
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 3 Then
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 3
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtCLimit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtConPers_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtConPers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtContactNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtOSBal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtOSBal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtTerms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
