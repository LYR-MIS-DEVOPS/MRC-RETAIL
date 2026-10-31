VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Begin VB.Form SPIncDisc 
   Caption         =   "Distribution of Discrepancy"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRemarks 
      Height          =   310
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtGood 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRow 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtNO 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Tag             =   "xyf"
      Top             =   1080
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxD 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   330
      BackColorBkg    =   -2147483633
      FormatString    =   "^PART NUMBER                |^DAMAGE |^LACKING|^EXCESS     |^WRONG PN"
   End
   Begin VB.Label Label1 
      Caption         =   "REMARKS:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1108
      Width           =   975
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   8655
      m_MDI_Height    =   2115
   End
End
Attribute VB_Name = "SPIncDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyRow As Integer


Private Sub cmdCancel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Cancel entry on Discrepancy?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbYes Then
        SPincoming.flxParts.TextMatrix(Val(txtRow.Text), 5) = SPincoming.flxParts.TextMatrix(Val(txtRow.Text), 4)
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    Dim TtlD As Integer, myRem As String
    
    With flxD
        TtlD = Val(.TextMatrix(1, 1)) + Val(.TextMatrix(1, 2)) + Val(.TextMatrix(1, 3)) + Val(.TextMatrix(1, 4))
        If TtlD + Val(txtGood.Text) <> Val(.TextMatrix(1, 5)) Then
            MsgBox "Invalid Total Qty. of Discrepancy!", vbExclamation + vbOKOnly, "(QTY=GOOD+DISCREPANCY)"
            Exit Sub
        End If
    End With
    With SPincoming.flxParts
        .TextMatrix(Val(txtRow.Text), 9) = flxD.TextMatrix(1, 1) 'DAMAGE
        .TextMatrix(Val(txtRow.Text), 10) = flxD.TextMatrix(1, 2) 'LACKING
        .TextMatrix(Val(txtRow.Text), 11) = flxD.TextMatrix(1, 3) 'EXCESS
        .TextMatrix(Val(txtRow.Text), 12) = flxD.TextMatrix(1, 4) 'WRONG PN
        .TextMatrix(Val(txtRow.Text), 15) = txtRemarks.Text  'REMARKS
        myRem = IIf(Val(flxD.TextMatrix(1, 1)) <> 0, "D(" & Val(flxD.TextMatrix(1, 1)) & ")", "")
        myRem = IIf(Val(flxD.TextMatrix(1, 2)) <> 0, IIf(myRem <> "", myRem & ", ", myRem) & "L(" & Val(flxD.TextMatrix(1, 2)) & ")", myRem)
        myRem = IIf(Val(flxD.TextMatrix(1, 3)) <> 0, IIf(myRem <> "", myRem & ", ", myRem) & "E(" & Val(flxD.TextMatrix(1, 3)) & ")", myRem)
        myRem = IIf(Val(flxD.TextMatrix(1, 4)) <> 0, IIf(myRem <> "", myRem & ", ", myRem) & "W(" & Val(flxD.TextMatrix(1, 4)) & ")", myRem)
        .TextMatrix(Val(txtRow.Text), 6) = myRem & " " & txtRemarks.Text
    End With
    Unload Me
End Sub

Private Sub flxD_Click()
    If flxD.Col = 0 Or flxD.Row = 0 Then Exit Sub
    With txtNO
        If flxD.TextMatrix(1, flxD.Col) <> "" Then txtNO.Tag = flxD.TextMatrix(1, flxD.Col)
        .Top = flxD.CellTop + flxD.Top
        .Left = flxD.CellLeft + flxD.Left
        .Width = flxD.CellWidth
        .Text = flxD.Text
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With

End Sub

Private Sub flxD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxD_Click
End Sub

Private Sub Form_Load()
    With SPincoming
        txtRow = MyRow
        txtGood = .flxParts.TextMatrix(MyRow, 5)
        flxD.TextMatrix(1, 0) = .flxParts.TextMatrix(MyRow, 1)
        flxD.TextMatrix(1, 1) = .flxParts.TextMatrix(MyRow, 9) 'Damage
        flxD.TextMatrix(1, 2) = .flxParts.TextMatrix(MyRow, 10) 'Lacking
        flxD.TextMatrix(1, 3) = .flxParts.TextMatrix(MyRow, 11) 'Excess
        flxD.TextMatrix(1, 4) = .flxParts.TextMatrix(MyRow, 12) 'wrong PN
        flxD.TextMatrix(1, 5) = .flxParts.TextMatrix(MyRow, 4) 'Qty
        txtRemarks.Text = .flxParts.TextMatrix(MyRow, 15)  'REMARKS

    End With

End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxD
        .ColWidth(0) = (.Width - 380) * 0.37
        .ColWidth(1) = (.Width - 380) * 0.16
        .ColWidth(2) = (.Width - 380) * 0.16
        .ColWidth(3) = (.Width - 380) * 0.16
        .ColWidth(4) = (.Width - 380) * 0.16
        .ColWidth(5) = 0
    End With

End Sub

Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtNO.Text = Empty
        flxD.TextMatrix(ActRow, 4) = txtNO.Tag
        txtNO.Tag = "skiplostfocus"
        txtNO.Visible = False
        flxD.SetFocus
    End If
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtNo_LostFocus
        flxD.SetFocus
    End If
End Sub

Private Sub txtNo_LostFocus()
    If txtNO.Tag = "skiplostfocus" Then
        txtNO.Tag = ""
        Exit Sub
    End If
    flxD.TextMatrix(1, flxD.Col) = txtNO.Text
    txtNO.Visible = False
End Sub

