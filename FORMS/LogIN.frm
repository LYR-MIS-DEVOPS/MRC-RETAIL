VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form LogIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log-In"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "LogIN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LogIN.frx":058A
   ScaleHeight     =   4035
   ScaleWidth      =   5550
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
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "LogIN.frx":566D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   2550
      TabIndex        =   4
      Top             =   3000
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "LogIN.frx":5C07
      cBack           =   -2147483633
   End
   Begin VB.Label lblLoginHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "==BRANCH NAME=="
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   270
      TabIndex        =   8
      Top             =   165
      Width           =   5055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your One stop Motorcycle Shop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   630
      TabIndex        =   7
      Top             =   945
      Width           =   4095
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
      Caption         =   "Your using the New MySQL 5.1 version"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
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
Dim rsMerInv As New ADODB.Recordset
Dim MerInvExist As Boolean

Private Sub cmdCancel_Click()
    cmdOK.Tag = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Call CheckID

End Sub

Private Sub Form_Load()
    'AniGIF1.Play
        
    Str1 = "Your using the New MySQL 5.1 version!!!! "
    Attempt = 0
    lblLoginHeader.Caption = SysBranch
    If InStr(SysBranch, "DEPOT") > 0 Then
        Label2.Caption = "Your Motorcycle Repair Solution"
    End If
    Call MnthlyMerInv

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
    Dim Qry As String, MyStat As String
    'Access
    'Qry = "SELECT PCUSER.* From PCUSER " & _
          "WHERE (((PCUSER.USERNAME)='" & txtUserName.Text & "') AND ((PCUSER.PASSWORD)='" & txtPW.Text & "'));"
    'mySQL
    Qry = "SELECT * From PCUSER " & _
          "WHERE PCUSER.USERNAME='" & txtUserName.Text & "' AND PCUSER.PASSWORD='" & txtPW.Text & "';"
    
    rsUser.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    'check for correct password
    If rsUser.RecordCount <= 0 Then
        MsgBox "Invalid User Name or Password!", vbExclamation + vbOKOnly, "Login"
        Attempt = Attempt + 1
        rsUser.Close
        If Attempt = 3 Then
            MsgBox "Sorry! You have 3 wrong attempts..." & _
                    vbCrLf & "This will terminate your session!", vbExclamation + vbOKOnly, "Login"
            End
        End If
        Exit Sub
    End If
    'Identify user type and accessible forms
    rsUser.MoveFirst
    'Access
    'If GetFieldValue("SELECT EMPLOYEES.EMPSTATUS From Employees WHERE (((EMPLOYEES.EMPID)=" & rsUser!EmpId & "));") = "In-Active" Then
    'mySQL
    If GetFieldValue("SELECT EMPLOYEES.EMPSTATUS From Employees WHERE EMPLOYEES.EMPID=" & rsUser!EMPID & ";") = "In-Active" Then
        
        MsgBox "Invalid User Name or Password!", vbExclamation + vbOKOnly, "Login"
        Attempt = Attempt + 1
        rsUser.Close
        If Attempt = 3 Then
            MsgBox "Sorry! You have 3 wrong attempts..." & _
                    vbCrLf & "This will terminate your session!", vbExclamation + vbOKOnly, "Login"
            End
        End If
        Exit Sub
    End If
    UserID = rsUser!EMPID
    UserType = rsUser!UserType
    
'    If rsUser!MOTOR <> 1 Then
'        MsgBox "Invalid User Name or Password!", vbExclamation + vbOKOnly, "Login"
'        Attempt = Attempt + 1
'        rsUser.Close
'        If Attempt = 3 Then
'            MsgBox "Sorry! You have 3 wrong attempts..." & _
'                    vbCrLf & "This will terminate your session!", vbExclamation + vbOKOnly, "Login"
'            End
'        End If
'        Exit Sub
'    End If
'    MotoRep.UpDown1.Tag = rsUser!DESIGNATION
'    If MotoRep.UpDown1.Tag = "System Admin." Or MotoRep.UpDown1.Tag = "Manager" Then
'        MotoRep.lblUser.Caption = MotoRep.UpDown1.Tag
'        MotoRep.UpDown1.Visible = True
'    Else
'        MotoRep.lblUser.Caption = MotoRep.UpDown1.Tag
'        MotoRep.UpDown1.Visible = False
'    End If
    'Call AccessForm
    Call UserAccess
    Call userPrivelege
    UserName = GetFieldValue("select concat(FIRSTNAME,' ',MIDDLENAME,' ',LASTNAME ) as empName from employees where empid= '" & UserID & "'")
    MotoRep.sBar.Panels(3).Text = UserName
    MotoRep.sBar.Panels(6).Text = "Welcome to " & SysBranch & "!"
    'MsgBox UserName
    rsUser.Close
    cmdOK.Tag = "OK"
    Unload Me
 End Sub


'Sub AccessForm()
'    Dim rsMyUser As New ADODB.Recordset
'    Dim MyMenu As Menu
'    Dim Qry As String
'    Dim i As Integer, j As Integer, N As Integer
'    Dim AccFrm(7) As String
'    Dim temp() As String
'    Dim mnuX As Menu
'
'    'Access
'    'Qry = "SELECT PCUSER.* From PCUSER WHERE (((PCUSER.EMPID)=" & UserID & "));"
'
'    'mySQL
'    Qry = "SELECT PCUSER.* From PCUSER WHERE PCUSER.EMPID=" & UserID & ";"
'
'    With rsMyUser
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        .MoveFirst
'        For i = 0 To 5
'            AccFrm(i) = IIf(IsNull(.Fields(i + 4)), "", .Fields(i + 4))
'        Next i
'        .Close
'    End With
'
'    For j = 0 To 5
'        temp = Split(AccFrm(j), ",")
'        Select Case j
'            Case 0
'                For i = LBound(temp) To UBound(temp)
'                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
'                    MotoRep.mSubPI(N).Enabled = True
'                Next i
'            Case 1
'                For i = LBound(temp) To UBound(temp)
'                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
'                    MotoRep.mSubSP(N).Enabled = True
'                    'If N = 4 Then MotoRep.BtnMain(0).Enabled = True
'                    'If N = 5 Then MotoRep.BtnMain(1).Enabled = True
'                Next i
'            Case 2
'                For i = LBound(temp) To UBound(temp)
'                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
'                    MotoRep.SubSRVC(N).Enabled = True
'                Next i
'            Case 3
'                For i = LBound(temp) To UBound(temp)
'                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
'                    MotoRep.mSubTR(N).Enabled = True
'                    If N = 5 Then MotoRep.BtnMain(4).Enabled = True
'                    If N = 7 Then MotoRep.BtnMain(1).Enabled = True
'                Next i
'            Case 4
'                For i = LBound(temp) To UBound(temp)
'                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
'                    MotoRep.mSubRep(N).Enabled = True
'                    If N = 1 Then
'                        MotoRep.BtnMain(2).Enabled = True
'                        MotoRep.BtnMain(3).Enabled = True
'                    End If
'                    If N = 2 Then MotoRep.BtnMain(5).Enabled = True
'                Next i
'            Case 5
'                For i = LBound(temp) To UBound(temp)
'                    N = (Val(Left$(temp(i), Len(temp(i)) - 1)))
'                    MotoRep.mSubTools(N).Enabled = True
'                Next i
'        End Select
'    Next j
'End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "\" Then KeyAscii = 0
    If KeyAscii = 13 Then Call CheckID
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Public Sub MnthlyMerInv()
    Qry = "SELECT last_day(curdate()-interval 1 month) AS lstdayOfmnth, curdate() AS currentDate, dayname(curdate()) AS dayName, day(curdate()) AS firstDay "
    rsMerInv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsMerInv.RecordCount = 1 Then
            rsMerInv.MoveFirst
                LstDayMnth = rsMerInv!lstdayofmnth
                CurrDate = rsMerInv!firstday
                'MsgBox LstDayMnth & " " & CurrDate, vbInformation, "Date"
        End If
    rsMerInv.Close
    
    Qry = "SELECT DISTINCT mnthlyinv.MONTHEND From mnthlyinv " & _
            "Where mnthlyinv.MONTHEND = '" & Format(LstDayMnth, "yyyy-mm-dd") & "' Group By MONTHEND"
    
    With rsMerInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 1 Then
            MerInvExist = True
        Else
            MerInvExist = False
        End If
        .Close
    End With
    
    
    If CurrDate < 5 And MerInvExist = False Then
        'MsgBox "okey"
            Qry = "SELECT spareparts.SPAREPARTID, spareparts.QTYONHAND, spareparts.COST, spareparts.STOREPRICE, spareparts.DlrPrice From spareparts where spareparts.QTYONHAND > '0'"
                With rsMerInv
                    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                        If .RecordCount > 0 Then
                            .MoveFirst
                            Do While Not .EOF
                               LYRcon.Execute "INSERT INTO mnthlyinv " & _
                                                "VALUES(" & Val(!SparepartID) & ", " & Val(!QtyOnHand) & ", " & _
                                                "" & Val(!Cost) & ", " & Val(!STOREPRICE) & ", " & Val(!DlrPrice) & ", '" & Format(LstDayMnth, "YYYY-MM-DD") & "')"
                            .MoveNext
                            Loop
                        End If
                    .Close
                End With
        'Else
        'MsgBox "DI SUCCESS!!!"
    End If
    
   
End Sub
