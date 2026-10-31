VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form InvDtlSalRpt 
   Caption         =   "Spare Parts Detailed Sales Report"
   ClientHeight    =   9150
   ClientLeft      =   225
   ClientTop       =   -2055
   ClientWidth     =   11490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   11490
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtOrder 
      Height          =   285
      Left            =   960
      TabIndex        =   21
      Top             =   9720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Left            =   960
      TabIndex        =   19
      Top             =   9360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   7575
      Left            =   0
      TabIndex        =   8
      Tag             =   "XY"
      Top             =   1560
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13361
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin CRVIEWERLibCtl.CRViewer vwrNew 
      Height          =   8415
      Left            =   0
      TabIndex        =   7
      Tag             =   "XY"
      Top             =   720
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
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   11295
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9300
         TabIndex        =   14
         Top             =   285
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   7560
         TabIndex        =   13
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
         Image           =   "InvDtlSalRpt.frx":0000
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   300
         Left            =   5400
         TabIndex        =   12
         Top             =   262
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
         Format          =   101318657
         CurrentDate     =   39289
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   300
         Left            =   2760
         TabIndex        =   10
         Top             =   262
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
         Format          =   101318657
         CurrentDate     =   39289
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8040
         TabIndex        =   15
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
         Image           =   "InvDtlSalRpt.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "FROM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   285
         Width           =   855
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Covered:"
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
         Left            =   -720
         TabIndex        =   5
         Top             =   285
         Width           =   2505
      End
      Begin VB.Label lblHeading 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Covered:"
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
         Left            =   -720
         TabIndex        =   6
         Top             =   315
         Width           =   2505
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
      Image           =   "InvDtlSalRpt.frx":0B34
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
      Caption         =   "&Filter Records"
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
      Image           =   "InvDtlSalRpt.frx":10CE
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
      Image           =   "InvDtlSalRpt.frx":1668
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
      Image           =   "InvDtlSalRpt.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   18
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
      Image           =   "InvDtlSalRpt.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Label lblRemarks 
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   9480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
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
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   450
      Width           =   2610
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11730
      m_MDI_Height    =   9720
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
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
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   2610
   End
   Begin VB.Menu nmuSales 
      Caption         =   "Sales Report"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "InvDtlSalRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer, Qry As String
Dim ActRow As Integer, fRow As Integer, fCol As Integer

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSP
        For i = fRow To .Rows - 1
            For j = fCol To 9
                If InStr(UCase$(.TextMatrix(i, j)), UCase$(txtFind.Text)) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 9:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 9, i + 1, i)
                    fCol = IIf(j = 9, 1, j + 1)
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
        
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.REFERENCE, SPITR.TRANSDATE, SPITR.REMARKS, SPITR.OUT, SPITR.PRICE, SPITR.SOURCE, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
          "FROM (SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID) INNER JOIN SUPPLIER ON SPAREPARTS.SUPPLIERCODE = SUPPLIER.SUPPLIERCODE " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & dtpFrom.Value & "# And (SPITR.TRANSDATE)<=#" & dtpTo.Value & "#) AND ((SPITR.TRANSTYPE)='OUT')) ORDER BY SUPPLIER.SUPPLIERCODE, SPAREPARTS.PARTNUMBER, SPITR.TRANSDATE;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.REFERENCE, SPITR.TRANSDATE,SPAREPARTS.QTYONHAND, SPITR.`OUT`, SPAREPARTS.LOCATION, SUPPLIER.SUPPLIERCODE " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID INNER JOIN SUPPLIER ON SPAREPARTS.SUPPLIERCODE = SUPPLIER.SUPPLIERCODE " & _
          "WHERE (((SPITR.TRANSDATE)>='" & Format(dtpFrom.Value, "yyyy-mm-dd") & "' And (SPITR.TRANSDATE)<='" & Format(dtpTo.Value, "yyyy-mm-dd") & "') AND ((SPITR.TRANSTYPE)='OUT')) ORDER BY SPAREPARTS.LOCATION, SPAREPARTS.PARTNUMBER, SPITR.TRANSDATE;"
    
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
End Sub

Public Sub ShowFilteredRecords()
    Dim SelFrom As String
    
    'SelFrom = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.REFERENCE, SPITR.TRANSDATE, SPITR.REMARKS, SPITR.OUT, SPITR.PRICE, SPITR.SOURCE, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
              "FROM (SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID) INNER JOIN SUPPLIER ON SPAREPARTS.SUPPLIERCODE = SUPPLIER.SUPPLIERCODE "
    'Qry = SelFrom & " " & txtCriteria.Text & " " & IIf(txtOrder.Text <> "", txtOrder.Text & ";", "ORDER BY SUPPLIER.SUPPLIERCODE, SPAREPARTS.PARTNUMBER, SPITR.TRANSDATE;")
    'mySQL
    SelFrom = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.REFERENCE, SPITR.TRANSDATE, SPITR.REMARKS, SPITR.`OUT`, SPITR.PRICE, SPITR.SOURCE, SPAREPARTS.LOCATION, SUPPLIER.SUPPLIERCODE " & _
              "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID INNER JOIN SUPPLIER ON SPAREPARTS.SUPPLIERCODE = SUPPLIER.SUPPLIERCODE "
    Qry = SelFrom & " " & txtCriteria.Text & " " & IIf(txtOrder.Text <> "", txtOrder.Text & ";", "ORDER BY SPAREPARTS.LOCATION, SPAREPARTS.PARTNUMBER, SPITR.TRANSDATE;")
    
    If rsRPT.State = adStateOpen Then rsRPT.Close
    
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsRPT.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 10
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsRPT
    flxSP.Visible = True
    Call RXResizer1_ResizeComplete
End Sub


Private Sub cmdTrans_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            'Call cmdRefresh_Click
            Call PrintPreView(False)
        Case 1  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
        Case 2  'Print Preview
            BtnIndx = 4
            tmrFloatRev.Enabled = True
            Call DisplayReport
            Call PrintPreView(True)
        Case 3  'Filter
            For i = 0 To 2
                 With SPfilter.cboLookIn(i)
                   .AddItem "PART NUMBER"
                   .AddItem "DESCRIPTION"
                   .AddItem "INVOICE NO."
                   .AddItem "INVOICE TYPE"
                   .AddItem "QUANTITY"
                   .AddItem "CUSTOMER"
                   .AddItem "SUPPLIER"
                 End With
            Next i
                                'SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.REFERENCE, SPITR.TRANSDATE, SPITR.REMARKS, SPITR.OUT, SPITR.PRICE, SPITR.SOURCE, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE
            SPfilter.fldNames = "SPAREPARTS.PARTNUMBER*SPAREPARTS.DESCRIPTION*SPITR.REFERENCE*SPITR.REMARKS*SPITR.OUT*SPITR.SOURCE*SPAREPARTS.SUPPLIERCODE"
            SPfilter.fldNdx = "4"
            SPfilter.fldDates = ""
            SPfilter.fldSupplier = 6
            SPfilter.MySource = "DETAILEDSALES"
            SPfilter.Show vbModal
            
        Case 4  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
    End Select
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    BtnIndx = 1
    tmrFloat.Enabled = True
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFrom.Value))
    Call cmdRefresh_Click
    Call PrintPreView(False)
    
    MotoRep.sBar.Panels(6).Text = "Detailed summary of Sold Spare Parts"
End Sub

Private Sub DisplayReport()
    Dim crSPsale As New RptDtldSales
    Dim MyDate As String
    
    MyDate = "Date Covered : " & Format(dtpFrom.Value, "MMM. dd, yyyy") & " to " & Format(dtpTo.Value, "MMM. dd, yyyy")
                        
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    crSPsale.DiscardSavedData
    crSPsale.Database.SetDataSource rsRPT
    crSPsale.EnableParameterPrompting = False
    crSPsale.ParameterFields(1).ClearCurrentValueAndRange
    crSPsale.ParameterFields(1).AddCurrentValue MyDate
    crSPsale.ParameterFields(2).ClearCurrentValueAndRange
    crSPsale.ParameterFields(2).AddCurrentValue lblRemarks.Caption
    crSPsale.ParameterFields(3).ClearCurrentValueAndRange
    crSPsale.ParameterFields(3).AddCurrentValue SysBranch
    crSPsale.ParameterFields(4).ClearCurrentValueAndRange
    crSPsale.ParameterFields(4).AddCurrentValue SysAddress
    vwrNew.ReportSource = crSPsale
    vwrNew.EnableExportButton = True
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
        If rsRPT.State = adStateOpen Then
            flxSP.Visible = True
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub RXResizer1_ResizeComplete()
    
    With flxSP
        .FormatString = "   |<PART NUMBER|<DESCRIPTION|^INVOICE NO.|^INV. DATE|^REMAINING QTY|^SALES QTY.|<LOCATION"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.15    'PN
        .ColWidth(2) = (.Width - 300) * 0.3     'DESC
        .ColWidth(3) = (.Width - 300) * 0.1     'INVNO
        .ColWidth(4) = (.Width - 300) * 0.1     'INV DATE
        .ColWidth(5) = (.Width - 300) * 0.1     'REMAINING QTY
        .ColWidth(6) = (.Width - 300) * 0.1     'SALES QTY
        .ColWidth(7) = (.Width - 300) * 0.15    'LOCATION
        .ColWidth(8) = 0 '(.Width - 300) * 0.15 'ID
        '.ColWidth(9) = 0 '(.Width - 300) * 0.15
        '.ColWidth(10) = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, 4) = Format$(.TextMatrix(i, 4), "mm/dd/yyyy")
            .TextMatrix(i, 7) = Format$(.TextMatrix(i, 7), "#,##0.00")
        Next i
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
    If KeyAscii = 13 Then
        Call cmdFind_Click
    End If
End Sub
