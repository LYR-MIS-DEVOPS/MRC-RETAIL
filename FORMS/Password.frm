VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3765
   Icon            =   "Password.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPW 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   435
      PasswordChar    =   "!"
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin lvButton.lvButtons_H cmdLocOK 
      Height          =   375
      Left            =   412
      TabIndex        =   1
      Top             =   960
      Width           =   1380
      _ExtentX        =   2434
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
      Image           =   "Password.frx":058A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdLocCancel 
      Height          =   375
      Left            =   1972
      TabIndex        =   2
      Top             =   960
      Width           =   1380
      _ExtentX        =   2434
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
      Image           =   "Password.frx":0B24
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MySource As String
Dim Atmp As Integer, PWok As Boolean

Private Sub cmdLocCancel_Click()
    Unload Me
End Sub

Private Sub cmdLocOK_Click()
    Dim Qry As String, myPW As String
    
    If MySource = "AUDITINV" Then
        'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.REMARKS)='PASSWORD') AND ((SETTINGS.DESCRIPTION)='INVENTORY AUDIT'));"
        'mySQL
        Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.REMARKS='PASSWORD' AND SETTINGS.DESCRIPTION='INVENTORY AUDIT';"
        
        myPW = GetFieldValue(Qry)
        If txtPW.Text = myPW Then
            PWok = True
            Unload Me
        Else
            Atmp = Atmp + 1
            MsgBox "Invalid Password!", vbCritical + vbOKOnly, "Audit"
        End If
        If Atmp = 3 Then
            AuditInvntry.cmdTrans(3).Tag = "INVALID PW"
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Atmp = 0
    PWok = False
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not PWok Then AuditInvntry.cmdTrans(3).Tag = "INVALID PW"
End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdLocOK_Click
End Sub
