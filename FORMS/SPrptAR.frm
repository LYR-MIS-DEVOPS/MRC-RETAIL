VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPrptAR 
   Caption         =   "Accounts Receivable Summary"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11490
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Tag             =   "XYF"
      Top             =   7800
      Width           =   11415
      Begin VB.Label lblTtlAmt 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL AMOUNT RECEIVABLE: 0.00"
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
         Left            =   7560
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   200
         Width           =   3735
      End
      Begin VB.Label lblRec 
         Caption         =   "RECORD X OF Y"
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
         Left            =   120
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   6615
      Left            =   0
      TabIndex        =   8
      Tag             =   "XY"
      Top             =   1200
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
      Height          =   7215
      Left            =   0
      TabIndex        =   7
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   525
      Width           =   11295
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9300
         TabIndex        =   11
         Top             =   285
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   225
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Image           =   "SPrptAR.frx":0000
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   300
         Left            =   960
         TabIndex        =   9
         Top             =   255
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20709377
         CurrentDate     =   39289
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8040
         TabIndex        =   12
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
         Image           =   "SPrptAR.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "As of:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   -600
         TabIndex        =   5
         Top             =   285
         Width           =   1500
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "As of:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   -600
         TabIndex        =   6
         Top             =   315
         Width           =   1500
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
      Image           =   "SPrptAR.frx":0B34
      Enabled         =   0   'False
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
      Image           =   "SPrptAR.frx":10CE
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
      Image           =   "SPrptAR.frx":1668
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
      Image           =   "SPrptAR.frx":1C02
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11610
      m_MDI_Height    =   8805
   End
   Begin VB.Menu nmuSales 
      Caption         =   "Sales Report"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "SPrptAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer, Qry As String
Dim fRow As Integer, fCol As Integer
Public ActRow As Integer

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSP
        For i = fRow To .Rows - 1
            For j = fCol To 10
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 10:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 10, i + 1, i)
                    fCol = IIf(j = 10, 1, j + 1)
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
    Dim i As Integer, Ttl As Currency
    
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CONTACTPERSON, ONACCTERMS.SPINVNO, ONACCTERMS.JOBORDERNO, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, [PAYSCHEME] & ' [' & [NPAYS] & ' Payments]' AS TERM, ONACCTERMS.BALANCE, ONACCTERMS.AMTDUE, ONACCTERMS.NXTBILL, ONACCTERMS.LASTUPDATE " & _
          "FROM CUSTOMER INNER JOIN ONACCTERMS ON CUSTOMER.CUSTOMERID = ONACCTERMS.CUSTOMERID WHERE (((ONACCTERMS.BALANCE)>0) AND ((ONACCTERMS.STATUS)='OPEN'));"

    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    
    If rsRPT.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsRPT
    flxSP.Visible = True
    With flxSP
        Ttl = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, 5) = Format$(.TextMatrix(i, 5), "mm/dd/yyyy")
            .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
            Ttl = Ttl + Val(.TextMatrix(i, 8))
            .TextMatrix(i, 8) = Format$(.TextMatrix(i, 8), "#,##0.00")
            .TextMatrix(i, 9) = Format$(.TextMatrix(i, 9), "#,##0.00")
            .TextMatrix(i, 10) = Format$(.TextMatrix(i, 10), "mm/dd/yyyy")
            .TextMatrix(i, 11) = Format$(.TextMatrix(i, 11), "mm/dd/yyyy")
        Next i
        lblTtlAmt.Caption = "TOTAL AMOUNT RECEIVABLE: " & Format(Ttl, "#,##0.00")
        If .Rows > 1 Then
            .Col = 1: .Row = 1
            .ColSel = 10:  .RowSel = 1
            ActRow = 1
            lblRec.Caption = "RECORD " & ActRow & " OF " & flxSP.Rows - 1
        End If
    End With
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call cmdRefresh_Click
            Call PrintPreView(False)
        Case 1  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
        Case 2  'Print Preview
            BtnIndx = 3
            tmrFloatRev.Enabled = True
            Call DisplayReport
            Call PrintPreView(True)
        Case 3  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
    End Select
End Sub


Private Sub flxSP_Click()
    ActRow = flxSP.Row
    lblRec.Caption = "RECORD " & ActRow & " OF " & flxSP.Rows - 1
End Sub

Private Sub flxSP_EnterCell()
    ActRow = flxSP.Row
    lblRec.Caption = "RECORD " & ActRow & " OF " & flxSP.Rows - 1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    BtnIndx = 1
    tmrFloat.Enabled = True
    dtpFrom.Value = Date
    Call cmdRefresh_Click
    Call PrintPreView(False)
    
    MotoRep.sBar.Panels(6).Text = "Summary of Accounts Receivables"
End Sub

Private Sub DisplayReport()
    Dim crSPinv As New RPTspMerchandise
    Dim MyDate As String
    
    MyDate = Format(dtpFrom.Value, "MMM. dd, yyyy")
        
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.REFERENCE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION, SPITR.SOURCE, SPAREPARTS.REMARKS " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON (SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID) AND (SPAREPARTS.REFERENCE = SPITR.REFERENCE) Where (((SPAREPARTS.QTYONHAND) <> 0)) " & _
          "ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.PARTNUMBER;"
    
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    crSPinv.DiscardSavedData
    crSPinv.Database.SetDataSource rsRPT
    crSPinv.EnableParameterPrompting = False
    crSPinv.ParameterFields(1).ClearCurrentValueAndRange
    crSPinv.ParameterFields(1).AddCurrentValue MyDate
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
    
    With flxSP
        .FormatString = "   |<CUSTOMER NAME|<EMPLOYER/MGR.|^INVOICE NO.|^JOB ORDER NO. |^ DATE|>INV. AMOUNT|<TERM|>BALANCE|>AMT. DUE |>DUE DATE |LAST UPDATE"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.15
        .ColWidth(2) = (.Width - 300) * 0.15
        .ColWidth(3) = (.Width - 300) * 0.08
        .ColWidth(4) = (.Width - 300) * 0.08
        .ColWidth(5) = (.Width - 300) * 0.08
        .ColWidth(6) = (.Width - 300) * 0.06
        .ColWidth(7) = (.Width - 300) * 0.15
        .ColWidth(8) = (.Width - 300) * 0.08
        .ColWidth(9) = (.Width - 300) * 0.08
        .ColWidth(10) = (.Width - 300) * 0.08
        .ColWidth(11) = (.Width - 300) * 0.08
    End With
    
    cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
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

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub
