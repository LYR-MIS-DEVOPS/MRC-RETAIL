VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPMerInvCost 
   Caption         =   "Spare Parts Merchandise Inventory"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11490
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   6615
      Left            =   0
      TabIndex        =   6
      Tag             =   "XY"
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11668
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin CRVIEWERLibCtl.CRViewer vwrNew 
      Height          =   7575
      Left            =   0
      TabIndex        =   5
      Tag             =   "XY"
      Top             =   600
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   525
      Width           =   11295
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9300
         TabIndex        =   7
         Top             =   285
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8040
         TabIndex        =   8
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
         Image           =   "SPMerInvCost.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label lblInvCost 
         Caption         =   "Total Inventory Value:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   240
         Width           =   4935
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
      TabIndex        =   0
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
      Image           =   "SPMerInvCost.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   1
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
      Image           =   "SPMerInvCost.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
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
      Image           =   "SPMerInvCost.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   3
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
      Image           =   "SPMerInvCost.frx":1668
      cBack           =   -2147483633
   End
   Begin VB.Label lblToolTip 
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Tag             =   "xyf"
      Top             =   0
      Width           =   2775
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11730
      m_MDI_Height    =   8760
   End
   Begin VB.Menu nmuSales 
      Caption         =   "Sales Report"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "SPMerInvCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer, Qry As String
Dim fRow As Integer, fCol As Integer

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSP
        For i = fRow To .Rows - 1
            For j = fCol To 7
                If InStr(.TextMatrix(i, j), UCase$(txtFind.Text)) > 0 Then
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

Private Sub cmdRefresh_Click()
    Dim InVal As Currency
    Dim InvQ As String
    
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.COST, [COST]*[QTYONHAND] AS TTLCOST, SPAREPARTS.LOCATION, SPAREPARTS.REMARKS " & _
          "From SPAREPARTS Where (((SPAREPARTS.QTYONHAND) <> 0)) ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.STOREPRICE as COST, SPAREPARTS.STOREPRICE*SPAREPARTS.QTYONHAND AS TTLCOST, SPAREPARTS.LOCATION, SPAREPARTS.REMARKS " & _
          "From SPAREPARTS Where (((SPAREPARTS.QTYONHAND) <> 0)) ORDER BY SPAREPARTS.LOCATION,SPAREPARTS.REMARKS, SPAREPARTS.PARTNUMBER;"
    
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsRPT.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 8
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsRPT
    ActRow = 1
    flxSP.Visible = True
    Call RXResizer1_ResizeComplete
    
    '***Access
    'InvQ = "SELECT Sum([COST]*[QTYONHAND]) AS TTLCOST From SPAREPARTS WHERE (((SPAREPARTS.QTYONHAND)<>0));"
    'mySQL
    InvQ = "SELECT Sum(SPAREPARTS.STOREPRICE*SPAREPARTS.QTYONHAND) AS TTLCOST From SPAREPARTS WHERE SPAREPARTS.QTYONHAND<>0;"
    
    InVal = Val(GetFieldValue(InvQ))
    lblInvCost.Caption = "TOTAL INVENTORY VALUE: " & Format(InVal, "#,##0.00")
    
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call cmdRefresh_Click
            Call PrintPreView(False)
        Case 1  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
        Case 2  'Print Preview
            BtnIndx = 3
            tmrFloatRev.Enabled = True
            Call DisplayReport
            Call PrintPreView(True)
        Case 3  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    lblToolTip.ToolTipText = MotoRep.Caption
    BtnIndx = 1
    tmrFloat.Enabled = True
    lblDate.Caption = "DATE: " & Format(Date, "Ddd, Mmmm dd, yyyy")
    Call cmdRefresh_Click
    Call PrintPreView(False)
    
    MotoRep.sBar.Panels(6).Text = "Spare Parts Merchandise Inventory"
End Sub

Private Sub DisplayReport()
    Dim crSPinv As New rptMerInvCost
    Dim MyDate As String
    
    'MyDate = Format(dtpFrom.Value, "MMM. dd, yyyy")
        
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.COST, [COST]*[QTYONHAND] AS TTLCOST, SPAREPARTS.LOCATION, SPAREPARTS.REMARKS " & _
          "From SPAREPARTS Where (((SPAREPARTS.QTYONHAND) <> 0)) ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.STOREPRICE as COST, SPAREPARTS.STOREPRICE*SPAREPARTS.QTYONHAND AS TTLCOST, SPAREPARTS.LOCATION, SPAREPARTS.REMARKS as CATEGORY " & _
          "From SPAREPARTS Where SPAREPARTS.QTYONHAND <> 0 ORDER BY SPAREPARTS.LOCATION, SPAREPARTS.REMARKS, SPAREPARTS.PARTNUMBER;"
    
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    crSPinv.DiscardSavedData
    crSPinv.ParameterFields(1).ClearCurrentValueAndRange
    crSPinv.ParameterFields(1).AddCurrentValue SysBranch
    crSPinv.ParameterFields(2).ClearCurrentValueAndRange
    crSPinv.ParameterFields(2).AddCurrentValue SysTel
    crSPinv.Database.SetDataSource rsRPT
    crSPinv.EnableParameterPrompting = False
    'crSPinv.ParameterFields(1).ClearCurrentValueAndRange
    'crSPinv.ParameterFields(1).AddCurrentValue MyDate
    crSPinv.PaperSource = crPRBinTractor
    crSPinv.PaperSize = crDefaultPaperSize
    vwrNew.ReportSource = crSPinv
    vwrNew.ViewReport
End Sub

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
        Frame1.Visible = False
        flxSP.Visible = False
        vwrNew.Visible = True
    Else
        vwrNew.Visible = False
        Frame1.Visible = True
        flxSP.Visible = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    With flxSP
        .FormatString = "   |<PART NUMBER|<DESCRIPTION|^QTY.|>SRP|>TOTAL|<LOCATION|<CATEGORY"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.15
        .ColWidth(2) = (.Width - 300) * 0.25
        .ColWidth(3) = (.Width - 300) * 0.1
        .ColWidth(4) = (.Width - 300) * 0.1
        .ColWidth(5) = (.Width - 300) * 0.1
        .ColWidth(6) = (.Width - 300) * 0.12
        .ColWidth(7) = (.Width - 300) * 0.12
        For i = 1 To .Rows - 1
            .TextMatrix(i, 4) = Format(.TextMatrix(i, 4), "#,##0.00")
            .TextMatrix(i, 5) = Format(.TextMatrix(i, 5), "#,##0.00")
        Next i
    End With
    
    'cmdFind.Height = cmdRefresh.Height
    'cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
    txtFind.Left = Frame1.Width - txtFind.Width - 50
    cmdFind.Left = txtFind.Left - (cmdFind.Width + 50)
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 4 Then
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
