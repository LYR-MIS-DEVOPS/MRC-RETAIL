VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmChoose 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3150
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4485
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3975
      Begin VB.CommandButton cmdBut3 
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton cmdBut1 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton cmdBut2 
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   3495
      End
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1320
      _ExtentX        =   2328
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
      Image           =   "frmChoose.frx":0000
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MySource As String

Private Sub cmdBut1_CLICK()
On Error GoTo er
    If MySource = "SELECTEDRECEIVE" Then
        Unload Me
        SPincoming.Show
        MotoRep.imgBG.Visible = False
        MotoRep.toolBarButton.Visible = False
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdBut2_Click()
On Error GoTo er
    If MySource = "SELECTEDRECEIVE" Then
        Unload Me
        SPbranchIncoming.Show
        MotoRep.imgBG.Visible = False
        MotoRep.toolBarButton.Visible = False
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdBut3_Click()
On Error GoTo er
    If MySource = "SELECTEDRECEIVE" Then
        Unload Me
        SPincomingWhAgdao.Show
        MotoRep.imgBG.Visible = False
        MotoRep.toolBarButton.Visible = False
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If SysBranchCategory = "3S" Then
        
        Frame1.Height = 2415
        Me.Height = 3045
    
    End If
    
    If MySource = "SELECTEDRECEIVE" Then
        Me.Caption = "RECEIVE"
        cmdBut1.Caption = "RECEIVE FROM &WAREHOUSE"
        cmdBut2.Caption = "RECEIVE FROM &BRANCH"
        cmdBut3.Caption = "RECEIVE FROM WAREHOUSE &DAVAO"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

