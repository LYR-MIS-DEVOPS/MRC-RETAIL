VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmInvCtrlBPO 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   8520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   8520
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000010&
      Height          =   285
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   240
      TabIndex        =   18
      Top             =   6960
      Width           =   10815
      Begin VB.TextBox txtApprovedBy 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   5175
      End
      Begin VB.TextBox txtRemarks 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox txtPreparedBy 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label Label12 
         Caption         =   "REMARKS:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "PREPARED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "ORDER NO:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "APPROVED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1935
      End
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   8640
      TabIndex        =   17
      Top             =   2880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "SAVE TO WAREHOUSE"
      CapAlign        =   2
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
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrder 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   8
      BackColorBkg    =   16777215
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   10815
      Begin VB.TextBox txtSupplier 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtOrderStatus 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtOrderType 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtOrderDate 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtOrderNo 
         BackColor       =   &H80000003&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H80000003&
         Height          =   645
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   "ORDER STATUS:"
         Height          =   255
         Left            =   6600
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "ORDER TYPE:"
         Height          =   255
         Left            =   6600
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "ORDER DATE:"
         Height          =   255
         Left            =   6600
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "ORDER NO:"
         Height          =   255
         Left            =   6600
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "ADDRESS:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "PURCHASE TO:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&New Order"
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
      Image           =   "frmInvCtrlBPO.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1678
      TabIndex        =   26
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit Order"
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
      Image           =   "frmInvCtrlBPO.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   27
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "frmInvCtrlBPO.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   28
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Order"
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
      Image           =   "frmInvCtrlBPO.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   29
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Order"
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
      Image           =   "frmInvCtrlBPO.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "frmInvCtrlBPO.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "frmInvCtrlBPO.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL"
      Height          =   255
      Left            =   8760
      TabIndex        =   24
      Top             =   6600
      Width           =   855
   End
End
Attribute VB_Name = "frmInvCtrlBPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrSupplier()   As String
Dim Qry             As String
Dim frmAction       As String
Dim btnIndex        As Integer

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAddress.SetFocus
End Sub

Private Sub cmdTrans_Click(Index As Integer)
'new(0), edit(1), summary(2), print(3), exit(4), save(5), cancel(6)

    Select Case Index
        Case 0 'new(0)
            frmAction = "NEW"
            Call lockField(False)
            Call btnShow(True)
            
            'Generate Order no
            Qry = "CALL sp_mrcretails.sp_getBPOReference('" & SysBranchID & "', '" & sysBranchCode & "')"
            If rsAnyUse.State <> 0 Then rsAnyUse.Close
            Set rsAnyUse = LYRcon.Execute(Qry)
            Call addToTextBox(rsAnyUse, txtOrderNo)
            '---------------------------------------
            
            Call defaultSupplierMainWarehouse
            txtOrderDate.Text = Format(ServerTime, "MM-DD-YYYY")
            txtOrderType.Text = "STOCK"
            txtOrderStatus.Text = "PENDING"
            txtPreparedBy.Text = UserName
            txtApprovedBy.Text = UCase(SysBranchOIC)
            
            
            txtSupplier.SetFocus
        Case 1 'edit(1)
            frmAction = "EDIT"
        Case 2 'summary(2)
            
        Case 3 'print(3)
        
        Case 4 'exit(4)
            Unload Me
        Case 5 'save(5)
        
        Case 6 'cancel(6)
            Call btnShow(False)
            Call clearText
    End Select

End Sub


Private Sub Form_Load()
    Call formatFlex
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Close
End Sub

Private Sub Timer1_Timer()
'    DoEvents
'    If rsAnyUse.State <> 0 Then rsAnyUse.Close
'    Call modSupplier.sp_supplier(rsAnyUse)
'    If rsAnyUse.RecordCount > 0 Then
'        Call modTools.displayCboObject(cboSupplier, rsAnyUse, arrSupplier, "SUPPLIERID", "SUPPLIERNAME")
'    End If
'    Timer1.Enabled = False
End Sub

Private Sub lockField(bln As Boolean)
'    txtSupplier.Locked = bln
'    txtAddress.Locked = bln
'    txtOrderNo.Locked = bln
'    txtOrderDate.Locked = bln
'    txtOrderType.Locked = bln
'    txtOrderStatus.Locked = bln
    txtRemarks.Locked = bln
'    txtApprovedBy.Locked = bln
'    txtPreparedBy.Locked = bln
End Sub

Private Sub clearText()
    'txtAddress.Text = ""
    txtApprovedBy.Text = ""
    txtOrderDate.Text = ""
    txtOrderNo.Text = ""
    txtOrderStatus.Text = ""
    txtOrderType.Text = ""
    txtPreparedBy.Text = ""
    txtRemarks.Text = ""
End Sub

Private Sub defaultSupplierMainWarehouse()
        
    Qry = "SELECT branch.BRANCHNAME, Branch.Address From Branch Where branch.BRANCHID = 29"
    If rsAnyUse.State <> 0 Then rsAnyUse.Close
    Set rsAnyUse = LYRcon.Execute(Qry)
    Call modDefaultSupplierMainWarehouse(rsAnyUse, txtSupplier, txtAddress)
        
End Sub

Private Sub formatFlex()

    With flxOrder
        .FormatString = " |<PARTNUMBER|<DESCRIPTION|^SOQ|^QTY|^PRICE|^TOTAL PRICE"
        .ColWidth(0) = 200
        .ColWidth(1) = .Width * 0.3
        .ColWidth(2) = .Width * 0.3
        .ColWidth(3) = .Width * 0.08
        .ColWidth(4) = .Width * 0.08
        .ColWidth(5) = .Width * 0.08
        .ColWidth(6) = .Width * 0.12
        .ColWidth(7) = 0
                
    End With
   
    
End Sub

Private Sub btnShow(btnEffect As Boolean)

If btnEffect Then
    For btnIndex = 0 To 4
        cmdTrans(btnIndex).Visible = False
    Next btnIndex
        cmdTrans(5).Left = cmdTrans(0).Left
        cmdTrans(5).Visible = True
        cmdTrans(6).Left = cmdTrans(5).Width
        cmdTrans(6).Visible = True
Else
    For btnIndex = 0 To 4
        cmdTrans(btnIndex).Visible = True
    Next btnIndex
        cmdTrans(5).Left = cmdTrans(6).Width
        cmdTrans(5).Visible = False
        cmdTrans(6).Left = cmdTrans(5).Width
        cmdTrans(6).Visible = False
End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOrderNo.SetFocus
End Sub

Private Sub txtPreparedBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtApprovedBy.SetFocus
End Sub

Private Sub txtOrderDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOrderType.SetFocus
End Sub

Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOrderDate.SetFocus
End Sub

Private Sub txtOrderStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtOrderType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOrderStatus.SetFocus
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPreparedBy.SetFocus
End Sub

Private Sub txtSupplier_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAddress.SetFocus
End Sub
