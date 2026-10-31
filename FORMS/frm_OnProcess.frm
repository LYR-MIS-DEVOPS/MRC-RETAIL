VERSION 5.00
Begin VB.Form frm_OnProcess 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   ForeColor       =   &H00404040&
   Icon            =   "frm_OnProcess.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgLoading 
      Height          =   1440
      Left            =   1560
      Picture         =   "frm_OnProcess.frx":57E2
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P L E A S E W A I T"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label lblOnProcess 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O N  P R O C E S S"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frm_OnProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RecordType As String
Public OwnerForm As String
 
Private Sub Form_Load()
On Error GoTo er
    imgLoading.Height = 1200
    LoadOnProcess = True
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error!"
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LoadOnProcess = False
    Set frm_OnProcess = Nothing
End Sub
