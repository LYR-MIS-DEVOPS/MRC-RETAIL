VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form LogIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log-In"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "LogIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LogIN.frx":058A
   ScaleHeight     =   4035
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5040
      Top             =   3120
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
      Top             =   2160
      Width           =   1815
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   405
      Left            =   480
      TabIndex        =   3
      Top             =   3120
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
      Image           =   "LogIN.frx":10CA9
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
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
      Image           =   "LogIN.frx":11243
      cBack           =   -2147483633
   End
   Begin VB.Label lblBranchname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "branch"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts Transfer"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   4575
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
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Your One Stop Motorcycle Shop!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3780
      Width           =   3975
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
      Top             =   2160
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

Private Sub cmdOK_Click()
    Dim rsUser As New ADODB.Recordset
    Dim Qry As String, myPos As String, Desig As String
    Dim UserFound As Boolean
    
    Qry = "SELECT PCUSER.* From PCUSER " & _
          "WHERE (((PCUSER.USERNAME)='" & txtUserName.Text & "') AND ((PCUSER.PASSWORD)='" & txtPW.Text & "'));"
    'mySQL
'    Qry = "SELECT employees.empid, employees.USERNAME, employees.PASSWORD, employees.usertype From employees " & _
'          "WHERE employees.USERNAME='" & txtUserName.Text & "' AND employees.PASSWORD='" & txtPW.Text & "';"
'
    With rsUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            UserID = !EmpId
            UserType = IIf(IsNull(!UserType), "", !UserType)
            Desig = IIf(IsNull(!Designation), "", UCase$(!Designation))

            If Desig = "SYSTEM ADMIN." Or Desig = "AUDITOR" Then
                ShowCost = True
            Else
                ShowCost = False
            End If
            'If !Receiver = True Then
            'mySQL
'            If !Receiver = 1 Then
'
'                UserFound = True
'            Else
'                UserFound = False
'                myPos = "Your User Account is not permitted to Access this transaction."
'            End If
            UserFound = True
        Else
            UserFound = False
            myPos = "Invalid User Name or Password!"
        End If
        .Close
    End With
    If Not UserFound Then
        MsgBox myPos, vbExclamation + vbOKOnly, "Order"
    Else
        cmdOK.Tag = "OK"
        DateStart = ""
        DateEnd = ""
        Unload Me
        ParTrans.Caption = ParTrans.Caption & " " & SysBranch & " v16.01.09"
        ParTrans.Show
    End If
End Sub

Private Sub Form_Load()
    'AniGIF1.Play
    Str1 = "Your One Stop Motorcycle Shop!!!! "
    Attempt = 0
   lblBranchname.Caption = SysBranch
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdOK.Tag <> "OK" Then End
End Sub



Private Sub Timer1_Timer()
    lbl.Caption = Str1
    Str1 = Right$(Str1, Len(Str1) - 1) & Left$(Str1, 1)
End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
