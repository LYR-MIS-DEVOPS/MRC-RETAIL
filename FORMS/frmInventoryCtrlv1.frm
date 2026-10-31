VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInventoryCtrlv1 
   Caption         =   "Inventory Control Summary"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtInvID 
      Height          =   375
      Left            =   7560
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   12720
      Top             =   0
   End
   Begin TabDlg.SSTab tabProcessing 
      Height          =   975
      Left            =   4560
      TabIndex        =   10
      Tag             =   "xyf"
      Top             =   3240
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1720
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Inventory Control"
      TabPicture(0)   =   "frmInventoryCtrlv1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProcess"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Image Image2 
         Height          =   720
         Left            =   240
         Picture         =   "frmInventoryCtrlv1.frx":001C
         Stretch         =   -1  'True
         Tag             =   "xyf"
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Tag             =   "xyf"
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "PLEASE WAIT...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Tag             =   "xyf"
         Top             =   360
         Width           =   2895
      End
   End
   Begin TabDlg.SSTab tabInvCtrl 
      Height          =   7095
      Left            =   120
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   480
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   12515
      _Version        =   393216
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Detailed"
      TabPicture(0)   =   "frmInventoryCtrlv1.frx":0A14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTotPN"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtTo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtFrom"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tabInvCtrlDtls"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Print Preview"
      TabPicture(1)   =   "frmInventoryCtrlv1.frx":0A30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabInvCtrlPreview"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Percentage"
      TabPicture(2)   =   "frmInventoryCtrlv1.frx":0A4C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "crInvCtrlGraph"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab tabInvCtrlPreview 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   15
         Tag             =   "xyf"
         Top             =   360
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   11668
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "MOVING"
         TabPicture(0)   =   "frmInventoryCtrlv1.frx":0A68
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "crInvCtrlMoving"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "NON MOVING"
         TabPicture(1)   =   "frmInventoryCtrlv1.frx":0A84
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "crInvCtrlNonMoving"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "UNDERSTOCKING"
         TabPicture(2)   =   "frmInventoryCtrlv1.frx":0AA0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "crInvCtrlUndrStocking"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin CRVIEWERLibCtl.CRViewer crInvCtrlNonMoving 
            Height          =   6135
            Left            =   120
            TabIndex        =   17
            Tag             =   "xyf"
            Top             =   360
            Width           =   13695
            DisplayGroupTree=   0   'False
            DisplayToolbar  =   -1  'True
            EnableGroupTree =   0   'False
            EnableNavigationControls=   -1  'True
            EnableStopButton=   -1  'True
            EnablePrintButton=   -1  'True
            EnableZoomControl=   -1  'True
            EnableCloseButton=   0   'False
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
         Begin CRVIEWERLibCtl.CRViewer crInvCtrlMoving 
            Height          =   6135
            Left            =   -74880
            TabIndex        =   16
            Tag             =   "xyf"
            Top             =   360
            Width           =   13695
            DisplayGroupTree=   0   'False
            DisplayToolbar  =   -1  'True
            EnableGroupTree =   0   'False
            EnableNavigationControls=   -1  'True
            EnableStopButton=   -1  'True
            EnablePrintButton=   -1  'True
            EnableZoomControl=   -1  'True
            EnableCloseButton=   0   'False
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
         Begin CRVIEWERLibCtl.CRViewer crInvCtrlUndrStocking 
            Height          =   6135
            Left            =   -74880
            TabIndex        =   18
            Tag             =   "xyf"
            Top             =   360
            Width           =   13695
            DisplayGroupTree=   0   'False
            DisplayToolbar  =   -1  'True
            EnableGroupTree =   0   'False
            EnableNavigationControls=   -1  'True
            EnableStopButton=   -1  'True
            EnablePrintButton=   -1  'True
            EnableZoomControl=   -1  'True
            EnableCloseButton=   0   'False
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
      End
      Begin TabDlg.SSTab tabInvCtrlDtls 
         Height          =   6255
         Left            =   120
         TabIndex        =   13
         Tag             =   "xyf"
         Top             =   720
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   11033
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "MOVING"
         TabPicture(0)   =   "frmInventoryCtrlv1.frx":0ABC
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "flxInvCtrlMoving"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "NON MOVING"
         TabPicture(1)   =   "frmInventoryCtrlv1.frx":0AD8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "flxInvCtrlNonMoving"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "UNDERSTOCKING"
         TabPicture(2)   =   "frmInventoryCtrlv1.frx":0AF4
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "flxInvCtrlUndrStocking"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInvCtrlMoving 
            Height          =   5775
            Left            =   -74880
            TabIndex        =   14
            Tag             =   "xyf"
            Top             =   360
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   10186
            _Version        =   393216
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInvCtrlNonMoving 
            Height          =   5775
            Left            =   -74880
            TabIndex        =   19
            Tag             =   "xyf"
            Top             =   360
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   10186
            _Version        =   393216
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInvCtrlUndrStocking 
            Height          =   5775
            Left            =   120
            TabIndex        =   20
            Tag             =   "xyf"
            Top             =   360
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   10186
            _Version        =   393216
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Tag             =   "xyf"
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   73662465
         CurrentDate     =   42331
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Tag             =   "xyf"
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   73662465
         CurrentDate     =   42331
      End
      Begin CRVIEWERLibCtl.CRViewer crInvCtrlGraph 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   9
         Tag             =   "xyf"
         Top             =   480
         Width           =   13935
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   0   'False
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
      Begin VB.Label Label1 
         Caption         =   "Date From:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "xyf"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Date To:"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Tag             =   "xyf"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTotPN 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   10800
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin lvButton.lvButtons_H cmdLoadData 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Load Data"
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
      Image           =   "frmInventoryCtrlv1.frx":0B10
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11520
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "frmInventoryCtrlv1.frx":10AA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSaveExcel 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Save Excel"
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
      Image           =   "frmInventoryCtrlv1.frx":1644
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   12000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14625
      m_MDI_Height    =   8220
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Menu mnuInvCtrl 
      Caption         =   "inventory control"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmInventoryCtrlv1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPrint                     As New Recordset
Dim timerCount                  As Integer
Dim invCtrlID                   As Integer
Dim dsrInvCtrlMoving            As New rptInvCtrlMoving
Dim dsrInvCtrlNonMoving         As New rptInvCtrlNonMoving
Dim dsrInvCtrlUnderStocking     As New rptInvCtrlUnderStocking
Dim dsrInvCtrlGraph             As New rptInvCtrlGraph

'Public Sub loadInventoryCtrl()
'
'On Error GoTo LocalErr
'
'        Qry = "call sp_mrcretails.sp_InvCtrlMain('" & SysBranchID & "')"
'        DoEvents
'        If rsAnyUse.State <> 0 Then rsAnyUse.Close
'        Set rsAnyUse = LYRcon.Execute(Qry)
'            lblTotPN = ""
'            lblTotPN = "TOTAL PARTNUMBER"
'            Call loadFlxRecord(rsAnyUse, flxInvCtrlMoving, lblTotPN)
'            DoEvents
'            'detailed
''            lblProcess.Caption = "Loading Preview Components ..."
''            dsrInvControl.DiscardSavedData
''            dsrInvControl.Database.SetDataSource rsAnyUse
''            crInvCtrl.ReportSource = dsrInvControl
''            crInvCtrl.ViewReport
'            lblProcess.Caption = "Loading Preview Components ..."
'            Call displayCRviewer(rsAnyUse, crInvCtrl, dsrInvControl)
'            DoEvents
'
'        'graph
'        lblProcess.Caption = "Loading Graph Components ..."
'        DoEvents
'        Qry = "call sp_mrcretails.sp_InventoryControlGraph('" & SysBranchID & "')"
'        If rsAnyUse.State <> 0 Then rsAnyUse.Close
'        Set rsAnyUse = LYRcon.Execute(Qry)
''            dsrInvCtrlGraph.DiscardSavedData
''            dsrInvCtrlGraph.Database.SetDataSource rsAnyUse
''            crInvCtrlGraph.ReportSource = dsrInvCtrlGraph
''            crInvCtrlGraph.ViewReport
'            Call displayCRviewer(rsAnyUse, crInvCtrlGraph, dsrInvCtrlGraph)
'            DoEvents
'
'            dtFrom.Value = flxInvCtrl.TextMatrix(2, 8)
'            Call RXResizer1_ResizeComplete
'
'
'
'        Exit Sub
'LocalErr:
'    MsgBox Err.Description, vbCritical, "Error"
'
'End Sub

Public Sub loadInventoryCtrl()
    On Error GoTo LocalErr
        Qry = "Call sp_mrcretails.sp_InvCtrlMain('" & SysBranchID & "')"
        DoEvents
            If rsAnyUse.State <> 0 Then rsAnyUse.Close
            Set rsAnyUse = LYRcon.Execute(Qry)
            If rsAnyUse.RecordCount > 0 Then
                rsAnyUse.MoveFirst
                txtInvID.Text = ""
                txtInvID.Text = rsAnyUse!InvCtrlMainID
            End If
            
            'Moving
            Qry = "Call sp_mrcretails.sp_InvCtrlRptMoving('" & txtInvID.Text & "')"
            DoEvents
            If rsAnyUse.State <> 0 Then rsAnyUse.Close
            Set rsAnyUse = LYRcon.Execute(Qry)
            Call loadFlxRecord(rsAnyUse, flxInvCtrlMoving)
            Call displayCRviewer(rsAnyUse, crInvCtrlMoving, dsrInvCtrlMoving)
            
            'Non moving
            Qry = "Call sp_mrcretails.sp_InvCtrlRptNonMoving('" & txtInvID.Text & "')"
            DoEvents
            If rsAnyUse.State <> 0 Then rsAnyUse.Close
            Set rsAnyUse = LYRcon.Execute(Qry)
            Call loadFlxRecord(rsAnyUse, flxInvCtrlNonMoving)
            Call displayCRviewer(rsAnyUse, crInvCtrlNonMoving, dsrInvCtrlNonMoving)
            
            'understocking
            Qry = "Call sp_mrcretails.sp_InvCtrlRptUnderstocking('" & txtInvID.Text & "')"
            DoEvents
            If rsAnyUse.State <> 0 Then rsAnyUse.Close
            Set rsAnyUse = LYRcon.Execute(Qry)
            Call loadFlxRecord(rsAnyUse, flxInvCtrlUndrStocking)
            Call displayCRviewer(rsAnyUse, crInvCtrlUndrStocking, dsrInvCtrlUnderStocking)
            
            'Graph
            Qry = "Call sp_mrcretails.sp_InvCtrlRptGraph('" & txtInvID.Text & "')"
            DoEvents
            If rsAnyUse.State <> 0 Then rsAnyUse.Close
            Set rsAnyUse = LYRcon.Execute(Qry)
            Call displayCRviewer(rsAnyUse, crInvCtrlGraph, dsrInvCtrlGraph)
    Exit Sub
            
LocalErr:
    MsgBox Err.Description, vbCritical, "Error Loading Procedure"
End Sub

Private Sub cmdExit_Click()
    Unload Me
    MotoRep.tbMain.Visible = True
End Sub

Private Sub cmdLoadData_Click()
    'flxInvCtrl.Rows = 2
    'flxInvCtrl.Cols = 2
    lblTotPN.Visible = True
    'flxInvCtrl.Clear
    tabProcessing.Visible = True
    lblProcess.Caption = "Loading Data Components ..."
    DoEvents
    Timer1.Enabled = True
End Sub

Private Sub cmdSaveExcel_Click()
    
    MsgBox "Not yet available", vbInformation, "Save Excel"
        
End Sub

Private Sub Form_Load()
    MotoRep.sBar.Panels(6).Text = frmInventoryCtrl.Caption
    tabInvCtrl.Tab = 0
    tabInvCtrlDtls = 0
    dtFrom.Value = ServerTime
    dtTo.Value = ServerTime
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxInvCtrlMoving
        .FormatString = "  |<Partnumber|^Qty|^6 Mo.Sales Qty|^Last Date In|^Last Date Out|<Location|<Category|<6mnthdate|<SPID"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.2
        .ColWidth(2) = (.Width - 300) * 0.08
        .ColWidth(3) = (.Width - 300) * 0.08
        .ColWidth(4) = (.Width - 300) * 0.08
        .ColWidth(5) = (.Width - 300) * 0.08
        .ColWidth(6) = (.Width - 300) * 0.2
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
    End With
    With flxInvCtrlNonMoving
        .FormatString = "  |<Partnumber|^Qty|^6 Mo.Sales Qty|^Last Date In|^Last Date Out|<Location|<Category|<6mnthdate|<SPID"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.2
        .ColWidth(2) = (.Width - 300) * 0.08
        .ColWidth(3) = (.Width - 300) * 0.08
        .ColWidth(4) = (.Width - 300) * 0.08
        .ColWidth(5) = (.Width - 300) * 0.08
        .ColWidth(6) = (.Width - 300) * 0.2
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
    End With
    With flxInvCtrlUndrStocking
        .FormatString = "  |<Partnumber|^Qty|^6 Mo.Sales Qty|^Last Date In|^Last Date Out|<Location|<Category|<6mnthdate|<SPID"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.2
        .ColWidth(2) = (.Width - 300) * 0.08
        .ColWidth(3) = (.Width - 300) * 0.08
        .ColWidth(4) = (.Width - 300) * 0.08
        .ColWidth(5) = (.Width - 300) * 0.08
        .ColWidth(6) = (.Width - 300) * 0.2
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
    End With
End Sub

Private Sub Timer1_Timer()
    Call loadInventoryCtrl
    DoEvents
    Timer1.Enabled = False
    tabProcessing.Visible = False
    DoEvents
    Call RXResizer1_ResizeComplete
End Sub
