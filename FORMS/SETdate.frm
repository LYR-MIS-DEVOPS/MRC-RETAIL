VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SETdate 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1350
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   15
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin MSComCtl2.MonthView SetMV 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   20709377
      CurrentDate     =   38903
   End
End
Attribute VB_Name = "SETdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MySource As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If MySource = "Cash Flow" Then
        'CashFlowCashier.CFdate.Value = SetMV.Value
        'Call CashFlowCashier.ViewCashFlow
        'Unload Me
    End If
    If MySource = "ArCvSumm" Then
        ARCVSumm.lblDate.Caption = SetMV.Value
        Unload Me
        Call ARCVSumm.DisplaySumm
    End If
    If MySource = "MGR CASHFLOW NEW" Then
        CashFlowMgr.CFdate.Value = SetMV.Value
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    SetMV.Value = Date
End Sub

Private Sub SetMV_DateClick(ByVal DateClicked As Date)
    cmdOK.SetFocus
End Sub
