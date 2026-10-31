VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form LogIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log-In"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   Icon            =   "LogIN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LogIN.frx":058A
   ScaleHeight     =   3465
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   1590
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4275
      Top             =   3420
   End
   Begin VB.TextBox txtPW 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2070
      Width           =   1815
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   405
      Left            =   1380
      TabIndex        =   3
      Top             =   2490
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   714
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
      Image           =   "LogIN.frx":3DFA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   2580
      TabIndex        =   4
      Top             =   2490
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   714
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
      Image           =   "LogIN.frx":4394
      cBack           =   -2147483633
   End
   Begin VB.Label lblSysBranch 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEMBRANCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   45
      TabIndex        =   8
      Top             =   135
      Width           =   4515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spareparts Receiver"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   495
      TabIndex        =   7
      Top             =   1050
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "A Member of LYR Group of Companies!!!!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   495
      TabIndex        =   5
      Top             =   3060
      Width           =   4065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password::"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   2070
      Width           =   1335
   End
End
Attribute VB_Name = "LogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Str1 As String
Dim Attempt As Integer
Private Sub cmdCancel_Click()
    cmdOK.Tag = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'If txtUserID.Text = "" Then
    '    MsgBox "Please enter your User ID!", vbCritical + vbOKOnly, "Partsman"
    '    Exit Sub
    'Else
        Call CheckID
    'End If
End Sub

Private Sub Form_Load()
    'AniGIF1.Play
    lblSysBranch.Caption = SysBranch
    Str1 = "A Member of LYR Group of Companies!!!! "
    Attempt = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdOK.Tag <> "OK" Then End
End Sub

Private Sub Timer1_Timer()
    lbl.Caption = Str1
    Str1 = Right$(Str1, Len(Str1) - 1) & Left$(Str1, 1)
End Sub

Private Sub CheckID()
    Dim rsUser As New ADODB.Recordset
    Dim Qry As String, myPos As String
    Dim UserFound As Boolean
    
    'Qry = "SELECT PCUSER.* From PCUSER " & _
          "WHERE (((PCUSER.USERNAME)='" & txtUserName.Text & "') AND ((PCUSER.PASSWORD)='" & txtPW.Text & "'));"
    'mySQL
    Qry = "SELECT PCUSER.* From PCUSER " & _
          "WHERE PCUSER.USERNAME='" & txtUserName.Text & "' AND PCUSER.PASSWORD='" & txtPW.Text & "';"
    
    With rsUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            UserID = !EMPID
            UserType = IIf(IsNull(!UserType), "", !UserType)
            myPos = GetPosition(UserID)
            If myPos = "Receiver" Or !Receiver = 1 Then
                UserFound = True
            Else
                UserFound = False
            End If
        Else
            UserFound = False
        End If
        .Close
    End With
    If Not UserFound Then
        MsgBox "Invalid User Name or Password!", vbExclamation + vbOKOnly, "Order"
    Else
        cmdOK.Tag = "OK"
        Unload Me
        SPincoming.Show
    End If

End Sub

Function GetPosition(MyID As Integer) As String
    Dim rsPos As New ADODB.Recordset
    Dim Qry As String

    'Qry = "SELECT EMPLOYEES.EMPID, EMPLOYEES.POSITION " & _
          "From EMPLOYEES WHERE (((EMPLOYEES.EMPID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT EMPLOYEES.EMPID, EMPLOYEES.POSITION " & _
          "From EMPLOYEES WHERE EMPLOYEES.EMPID=" & MyID & ";"
    
    With rsPos
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetPosition = IIf(IsNull(!Position), "", !Position)
        Else
            GetPosition = ""
        End If
        .Close
    End With
End Function

Sub AccessForm()
    Dim rsMyUser As New ADODB.Recordset
    Dim MyMenu As Menu
    Dim Qry As String
    Dim i As Integer, j As Integer, N As Integer
    Dim AccFrm(7) As String
    Dim temp() As String
    
    'Qry = "SELECT PCUSER.* From PCUSER WHERE (((PCUSER.EMPID)=" & UserID & "));"
    'mySQL
    Qry = "SELECT PCUSER.* From PCUSER WHERE PCUSER.EMPID=" & UserID & ";"
    
    With rsMyUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        .MoveFirst
        For i = 0 To 5
            AccFrm(i) = IIf(IsNull(.Fields(i + 4)), "", .Fields(i + 4))
        Next i
        .Close
    End With
    
    For j = 0 To 5
        temp = Split(AccFrm(j), ",")
        Select Case j
            Case 0
                For i = LBound(temp) To UBound(temp)
                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
                    MotoRep.mSubPI(N).Enabled = True
                Next i
            Case 1
                For i = LBound(temp) To UBound(temp)
                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
                    MotoRep.mSubSP(N).Enabled = True
                    If N = 4 Then MotoRep.BtnMain(0).Enabled = True
                    If N = 5 Then MotoRep.BtnMain(1).Enabled = True
                Next i
            Case 2
                For i = LBound(temp) To UBound(temp)
                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
                    MotoRep.SubSRVC(N).Enabled = True
                Next i
            Case 3
                For i = LBound(temp) To UBound(temp)
                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
                    MotoRep.mSubTR(N).Enabled = True
                    If N = 5 Then MotoRep.BtnMain(4).Enabled = True
                Next i
            Case 4
                For i = LBound(temp) To UBound(temp)
                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
                    MotoRep.mSubRep(N).Enabled = True
                    If N = 1 Then
                        MotoRep.BtnMain(2).Enabled = True
                        MotoRep.BtnMain(3).Enabled = True
                    End If
                    If N = 2 Then MotoRep.BtnMain(5).Enabled = True
                Next i
            Case 5
                For i = LBound(temp) To UBound(temp)
                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
                    MotoRep.mSubTools(N).Enabled = True
                Next i
        End Select
    Next j
End Sub


Private Sub txtPW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call CheckID
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
