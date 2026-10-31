VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmViewDiscrepancy 
   Caption         =   "View Discrepancy"
   ClientHeight    =   8625
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   10800
      TabIndex        =   13
      Text            =   "HoldBackColor"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Tag             =   "XY"
      Top             =   7560
      Width           =   10455
      Begin VB.TextBox txtPrcssBy 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   360
         Width           =   6855
      End
      Begin VB.Label lblPrcssBy 
         Caption         =   "PROCESS BY:"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   120
      TabIndex        =   6
      Tag             =   "XY"
      Top             =   2160
      Width           =   10455
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDscrpncy 
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8705
         _Version        =   393216
         Cols            =   6
         BackColorBkg    =   8421504
         FormatString    =   "|^NO|^PARTNUMBER|^DESCRIPTION|^REMARKS|^LOCATION"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin MSMask.MaskEdBox txtDISDate 
      Height          =   360
      Left            =   8640
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   600
      Width           =   10455
      Begin VB.TextBox txtBase_ID 
         Height          =   375
         Left            =   -120
         TabIndex        =   14
         Top             =   9020
         Width           =   735
      End
      Begin VB.ComboBox cboBaseLine 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtBranch 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Tag             =   "X"
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE:"
         Height          =   255
         Left            =   7800
         TabIndex        =   5
         Tag             =   "XY"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblBaseLine 
         Caption         =   "BASE LINE:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Tag             =   "XY"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblBranch 
         Caption         =   "BRANCH:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Tag             =   "XY"
         Top             =   240
         Width           =   735
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   661
      Caption         =   "&Print"
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
      Image           =   "frmViewDiscrepancy.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   1620
      TabIndex        =   16
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "frmViewDiscrepancy.frx":059A
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   4440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   10905
      m_MDI_Height    =   9195
   End
End
Attribute VB_Name = "frmViewDiscrepancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CID As Long

Dim arrPOSCODE() As String, arrCCODE() As String, arrValidation() As String
Dim tmpRs  As New ADODB.Recordset
Dim mainRs As New ADODB.Recordset
Dim cboRs As New ADODB.Recordset


Dim Ans As String
Dim Qry As String
Dim strPartial As String, strTotal As String
Dim formAction As String, myTicketNo As String
Dim myFile As String, POSCODE As String, CCODE As String

Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer
Dim EMPID As Long

Dim StatSrch As Boolean
Dim PrtBln As Boolean, PrtSrch As Boolean, ProperExit As Boolean, myUpload As Boolean
Private Sub cboBaseLine_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cboBaseLine
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBaseLine-Change()")
End Sub
Private Sub CboDisplay()
On Error GoTo er
    
    'Get Data from database
    DoEvents
    Qry = "Select base_id, generated_baseline from inv_baseline group by base_id"
    If cboRs.State = adStateOpen Then cboRs.Close
    cboRs.Open Qry, LYRcon, adOpenStatic, adLockReadOnly
    If cboRs.State = 1 Then
       Call modObjTools.addToListObjectWithArrayGivenField(cboBaseLine, cboRs, arrPOSCODE, "base_id", "generated_baseline")
    End If
    If cboRs.State <> 0 Then cboRs.Close
    
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBaseLine()")
End Sub


Private Sub cboBaseLine_Click()
    txtBase_ID.Text = arrPOSCODE(cboBaseLine.ListIndex, 0)
    Call DisplayDiscrepancy
End Sub

Private Sub cboBaseLine_KeyDown(KeyCode As Integer, Shift As Integer)
 'If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then PmanSrch = True
End Sub

Private Sub cboBaseLine_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtDISDate.SetFocus
End Sub

Private Sub DisplayDiscrepancy()
    Dim Qry As String
    
    If Val(txtBase_ID.Text) = 0 Then Exit Sub
    'Qry = "SELECT BRANCH.ADDRESS From BRANCH WHERE (((BRANCH.BRANCHID)=" & Val(txtBranchID.Text) & "));"
    'mySQL
    
    Qry = "call sp_mrcretails.sp_ViewDiscrepancy('" & sysBranchCode & "', '" & Val(txtBase_ID.Text) & "')"
    If mainRs.State = adStateOpen Then mainRs.Close
    mainRs.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If mainRs.RecordCount = 0 Then
        flxDscrpncy.Clear
        flxDscrpncy.Rows = 2
        Call RXResizer1_ResizeComplete
        Exit Sub
    Else
        With mainRs
            .MoveFirst
            Call RXResizer1_ResizeComplete
            flxDscrpncy.Rows = .RecordCount + 1
            If .RecordCount > 0 Then
                For i = 1 To flxDscrpncy.Rows - 1
                    flxDscrpncy.TextMatrix(i, 1) = i
                    flxDscrpncy.TextMatrix(i, 2) = !PartNumber
                    flxDscrpncy.TextMatrix(i, 3) = !Description
                    flxDscrpncy.TextMatrix(i, 4) = !Remarks
                    flxDscrpncy.TextMatrix(i, 5) = !Location
                    'txtPrcssBy = IIf(IsNull(!PreparedBy), "", !PreparedBy)
                    
                    .MoveNext
                Next i
            End If
        End With
    End If
End Sub

'Private Sub cmdTrans_Click(Index As Integer)
'If Index = 0 Then 'print
'
'
'ElseIf Index = 1 Then 'exit
'    Call UnloadMe
'    Set frmViewDiscrepancy = Nothing
'End If
'End Sub

'Private Sub flxDscrpncy_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then Call flxDscrpncy_Click
'End Sub
Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    
    'frmHeader
    txtDISDate.Enabled = Not DataEntry
    txtPrcssBy.Locked = DataEntry
    'frmBody
    txtBranch.Locked = DataEntry

    If txtBranch.Locked Then
        'frmHeader
        txtPrcssBy.BackColor = txtHoldGrayBackColor.BackColor
        txtBranch.BackColor = txtHoldGrayBackColor.BackColor
        txtDISDate.BackColor = txtHoldGrayBackColor.BackColor
    Else
        'frmHeader
        txtPrcssBy.Locked = True
        txtBranch.Locked = True
        txtDISDate.Enabled = False
    End If
End Sub
Private Sub Form_Load()
    Call LockData(False)
    txtDISDate.Text = Format$(Now, "mm/dd/yyyy")
    txtBranch.Text = SysBranch
    Call CboDisplay
End Sub

Sub clearform()

    flxDscrpncy.Rows = 1
    flxDscrpncy.Rows = 2
    'Frame 1
    txtBranch.Text = SysBranch
    txtPrcssBy.Text = ""
    txtDISDate.Text = Format$(Now, "mm/dd/yyyy")
End Sub


Private Sub RXResizer1_ResizeComplete()
    On Error GoTo Err
    Dim i As Integer, j As Integer
    Dim aCol As Integer
    
    'lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    'lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    With flxDscrpncy
        
        For aCol = 0 To 2
           .ColAlignment(aCol) = 2
        Next aCol
        
         For aCol = 3 To 5
           .ColAlignment(aCol) = 4
        Next aCol
        .ColAlignment(5) = 4
        .TextMatrix(0, 1) = "NUMBER"
        .TextMatrix(0, 2) = "PARTNUMBER"
        .TextMatrix(0, 3) = "DESCRIPTION"
        .TextMatrix(0, 4) = "REMARKS"
        .TextMatrix(0, 5) = "LOCATION"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - 280) * 0.05            'NO
        .ColWidth(2) = (.Width - 280) * 0.3             'PARTNUMBER
        .ColWidth(3) = (.Width - 280) * 0.3              'DESCRIPTION
        .ColWidth(4) = (.Width - 280) * 0.13        'REMARKS
        .ColWidth(5) = (.Width - 280) * 0.1       'location

    End With

Err:

End Sub

