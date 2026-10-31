VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form AuditInvSum 
   Caption         =   "Inventory Audit Summary"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11520
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxInvCancel 
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "REFERENCE NO.|^REF. TYPE"
   End
   Begin MSFlexGridLib.MSFlexGrid flxAud 
      Height          =   6975
      Left            =   0
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   1200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "    |^  BATCH NO.    |^ DATE START    |^DATE END       |< STATUS           |^  SUB REPORTS       |<NAME OF AUDITOR         "
   End
   Begin CRVIEWERLibCtl.CRViewer vwrDaily 
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Tag             =   "XY"
      Top             =   480
      Width           =   11415
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   400
      Width           =   11295
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9240
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   285
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   7920
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Find Next"
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
         Image           =   "AuditInvSumm.frx":0000
         cBack           =   -2147483633
      End
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   8280
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   7800
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   8280
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   7800
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Report"
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
      Image           =   "AuditInvSumm.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Report"
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
      Image           =   "AuditInvSumm.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Grid View"
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
      Image           =   "AuditInvSumm.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Report"
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
      Image           =   "AuditInvSumm.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "AuditInvSumm.frx":1C02
      cBack           =   -2147483633
   End
   Begin VB.Label lblP 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11760
      m_MDI_Height    =   8745
   End
   Begin VB.Menu mnuDaily 
      Caption         =   "Daily Cash"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel O.R."
      End
      Begin VB.Menu mnuAttach 
         Caption         =   "Show Attachment"
      End
   End
End
Attribute VB_Name = "AuditInvSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Dim MQry As String

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxAdj
        For i = fRow To .Rows - 1
            For j = fCol To 7
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 7:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 7, i + 1, i)
                    fCol = IIf(j = 7, 1, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
        Case 1  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
        Case 2
            Call AuditInvntry.ShowRecord(Val(flxAud.TextMatrix(ActRow, 7)))
            Unload Me
        Case 3  'Print Preview
            BtnIndx = 4
            tmrFloatRev.Enabled = True
            Call DisplayReport
            Call PrintPreView(True)
        Case 4  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
    End Select
End Sub

Private Sub RefreshRecord()
    Dim i As Integer
    
    MQry = "SELECT AUDITINVEN.BATCHNO, AUDITINVEN.DATESTART, AUDITINVEN.DATEEND, AUDITINVEN.STATUS, AUDITINVEN.BATCHMEMBER, AUDITINVEN.AUDITOR, AUDITINVEN.BATCHID From AUDITINVEN " & _
           "ORDER BY AUDITINVEN.HOMEID DESC , AUDITINVEN.BATCHID;"
    
    If rsRPT.State = adStateOpen Then rsRPT.Close
    With rsRPT
        .Open MQry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxAud.Rows = 1
            Do While Not .EOF
                flxAud.AddItem vbTab & !BatchNo & vbTab & Format(!DateStart, "MM/DD/YYYY") & vbTab & IIf(IsNull(!DateEnd), "", Format(!DateEnd, "MM/DD/YYYY")) & vbTab & !STATUS & vbTab & IIf(IsNull(!BatchMember), 0, !BatchMember) & vbTab & !Auditor & vbTab & !BATCHID
                .MoveNext
            Loop
        Else
            flxAud.Rows = 1
            flxAud.Rows = 2
        End If
        .Close
    End With
End Sub


Private Sub flxAud_Click()
    If flxAud.Col = 0 And flxAud.Row > 0 Then
        ActRow = flxAud.Row
        Call flxAud_EnterCell
    End If
End Sub

Private Sub flxAud_EnterCell()
    ActRow = flxAud.Row
    With flxAud
        '.Row = ActRow: .Col = 0
        '.RowSel = ActRow: .ColSel = .Cols - 1
        If Not .RowIsVisible(.Row) Then .TopRow = .Row
    End With
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    BtnIndx = 1
    tmrFloat.Enabled = True
    Call PrintPreView(False)
    Call RefreshRecord
    ActRow = 1
    MotoRep.sBar.Panels(6).Text = "Invetory Audit Summary"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
End Sub

Private Sub DisplayReport()
    'Dim crDaily As New rptDailyCash
            
    'If rsRPT.State = adStateOpen Then rsRPT.Close
    'rsRPT.Open MQry, LYRcon, adOpenKeyset, adLockOptimistic
    'crDaily.DiscardSavedData
    'crDaily.Database.SetDataSource rsRPT
    'vwrDaily.ReportSource = crDaily
    'vwrDaily.ViewReport
End Sub

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
        Frame1.Visible = False
        flxAud.Visible = False
        vwrDaily.Visible = True
    Else
        vwrDaily.Visible = False
        Frame1.Visible = True
        flxAud.Visible = True
    End If
End Sub


Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    With flxAud
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - 500) * 0.13
        .ColWidth(2) = (.Width - 500) * 0.16
        .ColWidth(3) = (.Width - 500) * 0.16
        .ColWidth(4) = (.Width - 500) * 0.16
        .ColWidth(5) = (.Width - 500) * 0.13
        .ColWidth(6) = (.Width - 500) * 0.26
        .ColWidth(7) = 0
    End With
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 5 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 2 Then
            tmrDropRev.Enabled = False
            'cmdTrans(0).SetFocus
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 2 'Index of Button view grid
            cmdTrans(2).Top = -375
            cmdTrans(2).Left = 0
            cmdTrans(2).Visible = True
            cmdTrans(3).Top = -375
            cmdTrans(3).Left = cmdTrans(2).Width
            cmdTrans(3).Visible = True
            cmdTrans(4).Top = -375
            cmdTrans(4).Left = cmdTrans(2).Width * 2
            cmdTrans(4).Visible = True
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 1 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            cmdTrans(0).Visible = True
            cmdTrans(1).Visible = True
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

