VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrans 
      Caption         =   "View History"
      Height          =   615
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "View Discrepancy"
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
 frmViewDiscrepancy.Show
 
 
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    If Index = 0 Then 'View Discrepancy
        Call frmViewDiscrepancy.Show
    ElseIf Index = 1 Then
        Call frmViewHistory.Show
    End If
End Sub
