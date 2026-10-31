VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MonthlySalesAdjBreakDown 
   Caption         =   "Monthly Sales Report - Break Down"
   ClientHeight    =   9750
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14400
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   14400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtsp2 
      Height          =   375
      Left            =   5880
      TabIndex        =   41
      Top             =   11040
      Width           =   1935
   End
   Begin VB.TextBox txtsp 
      Height          =   375
      Left            =   3360
      TabIndex        =   40
      Top             =   11040
      Width           =   1935
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5160
      Top             =   0
   End
   Begin VB.Frame frmwait 
      Height          =   1215
      Left            =   3240
      TabIndex        =   38
      Tag             =   "XYF"
      Top             =   6240
      Visible         =   0   'False
      Width           =   8175
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait While Processing Sales Distribution"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1440
         TabIndex        =   39
         Top             =   480
         Width           =   4995
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   60
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   840
      Width           =   14295
      Begin VB.CommandButton Command2 
         Caption         =   "&Print Preview"
         Height          =   525
         Left            =   12720
         Picture         =   "MonthlySalesAdjBreakdown.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   210
         Width           =   1395
      End
      Begin VB.CheckBox chkSel 
         Caption         =   "SPAREPARTS"
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1485
      End
      Begin VB.ComboBox cboCutOff 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdRep 
         Caption         =   "Reprocess"
         Height          =   525
         Left            =   11280
         Picture         =   "MonthlySalesAdjBreakdown.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "For the Month of:"
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
         Left            =   1770
         TabIndex        =   30
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label lblMoColl 
         Alignment       =   2  'Center
         Caption         =   "Total Collection for the Month"
         Height          =   255
         Left            =   5550
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   390
         Width           =   5655
      End
   End
   Begin TabDlg.SSTab stabSales 
      Height          =   6345
      Left            =   60
      TabIndex        =   10
      Tag             =   "XYF"
      Top             =   3090
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   11192
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SPAREPARTS"
      TabPicture(0)   =   "MonthlySalesAdjBreakdown.frx":0B14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxHead"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "flxSP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "SERVICE CENTER"
      TabPicture(1)   =   "MonthlySalesAdjBreakdown.frx":0B30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxHead2"
      Tab(1).Control(1)=   "flxSP2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "MonthlySalesAdjBreakdown.frx":0B4C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSHFlexGrid1"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "PRINT PREVIEW"
      TabPicture(3)   =   "MonthlySalesAdjBreakdown.frx":0B68
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CRViewer1"
      Tab(3).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4935
         Left            =   -74940
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   1350
         Width           =   14115
         _ExtentX        =   24897
         _ExtentY        =   8705
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame3 
         Height          =   885
         Left            =   -74940
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   420
         Width           =   14115
         Begin VB.CommandButton Command9 
            Appearance      =   0  'Flat
            Caption         =   "&Execute"
            Height          =   525
            Left            =   7560
            Picture         =   "MonthlySalesAdjBreakdown.frx":0B84
            Style           =   1  'Graphical
            TabIndex        =   24
            Tag             =   "XYF"
            Top             =   210
            Width           =   915
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   5370
            TabIndex        =   23
            Tag             =   "XYF"
            Top             =   300
            Width           =   2145
         End
         Begin VB.TextBox Text1 
            Height          =   345
            Left            =   12030
            TabIndex        =   19
            Tag             =   "XYF"
            Top             =   270
            Width           =   1965
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Find &Next"
            Height          =   495
            Left            =   10620
            Picture         =   "MonthlySalesAdjBreakdown.frx":110E
            Style           =   1  'Graphical
            TabIndex        =   18
            Tag             =   "XYF"
            Top             =   210
            Width           =   1395
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Print Preview"
            Height          =   525
            Left            =   8550
            Picture         =   "MonthlySalesAdjBreakdown.frx":1698
            Style           =   1  'Graphical
            TabIndex        =   16
            Tag             =   "XYF"
            Top             =   210
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1800
            TabIndex        =   15
            Tag             =   "XYF"
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "SALES TYPE:"
            Height          =   195
            Left            =   4290
            TabIndex        =   22
            Tag             =   "XYF"
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "CUSTOMER NAME:"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Tag             =   "XYF"
            Top             =   360
            Width           =   1470
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxSP 
         Height          =   5145
         Left            =   90
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   1080
         Width           =   14085
         _ExtentX        =   24844
         _ExtentY        =   9075
         _Version        =   393216
         Rows            =   1
         Cols            =   15
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
         BorderStyle     =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxHead 
         Height          =   705
         Left            =   90
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   450
         Width           =   14025
         _ExtentX        =   24739
         _ExtentY        =   1244
         _Version        =   393216
         Cols            =   15
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxHead2 
         Height          =   615
         Left            =   -74910
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   450
         Width           =   14115
         _ExtentX        =   24897
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   11
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxSP2 
         Height          =   5145
         Left            =   -74910
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   1080
         Width           =   14115
         _ExtentX        =   24897
         _ExtentY        =   9075
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
         BorderStyle     =   0
      End
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   5745
         Left            =   -74880
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   450
         Width           =   14115
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
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   9450
      Visible         =   0   'False
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   0
   End
   Begin VB.TextBox txtYr 
      Height          =   285
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPPro 
      Height          =   615
      Left            =   840
      TabIndex        =   7
      Top             =   9840
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      FixedCols       =   0
      FormatString    =   "DATE |>WALKIN|>WHLSALE|>RBLTPLAN|>LGCMKTG|>LGCFIN|>WHLACCT|>EMPACCT|>CORTACCT|>RETURN|>AdjDebit|>AdjCredit|>SOTHERS|>SPTotal"
   End
   Begin VB.Frame frmDay 
      Caption         =   "Sales For the Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1275
      Left            =   60
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   1800
      Width           =   14295
      Begin MSFlexGridLib.MSFlexGrid flxSVRDay 
         Height          =   945
         Left            =   75
         TabIndex        =   34
         Tag             =   "XYF"
         Top             =   285
         Visible         =   0   'False
         Width           =   14115
         _ExtentX        =   24897
         _ExtentY        =   1667
         _Version        =   393216
         Rows            =   3
         Cols            =   10
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         BackColor       =   12648447
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxSPDay 
         Height          =   915
         Left            =   90
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   300
         Visible         =   0   'False
         Width           =   14025
         _ExtentX        =   24739
         _ExtentY        =   1614
         _Version        =   393216
         Rows            =   3
         Cols            =   15
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         BackColor       =   12648447
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Left            =   0
      TabIndex        =   36
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
      Image           =   "MonthlySalesAdjBreakdown.frx":1C22
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPSrv 
      Height          =   615
      Left            =   840
      TabIndex        =   37
      Top             =   10440
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      FormatString    =   "DATE|>WALKIN|>LGCMKTG|>LGCFIN|>EMPACCT|>CORTACCT|>AdjDebit|>AdjCredit|>SOTHERS|>SVRTOTAL"
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
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
      Left            =   13680
      TabIndex        =   8
      Tag             =   "XY"
      Top             =   9450
      Width           =   615
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   -180
      Top             =   9990
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14520
      m_MDI_Height    =   10260
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
      Left            =   120
      TabIndex        =   0
      Top             =   570
      Width           =   2610
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
      Left            =   150
      TabIndex        =   1
      Top             =   570
      Width           =   2610
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SALES BREAKDOWN"
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
      Index           =   2
      Left            =   9690
      TabIndex        =   2
      Tag             =   "XYF"
      Top             =   600
      Width           =   4500
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SALES BREAKDOWN"
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
      Left            =   9735
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   600
      Width           =   4500
   End
   Begin VB.Menu mnuINV 
      Caption         =   "SP Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuProcess 
         Caption         =   "Verify Sales"
      End
   End
End
Attribute VB_Name = "MonthlySalesAdjBreakDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMSALESADJ As New ADODB.Recordset
Dim ActRow As Integer
Dim ReProc As Boolean

Private Sub cboCutOff_Click()
    Dim MoYr() As String
   
    If cboCutOff.ListIndex = -1 Or cboCutOff.Text = "" Then Exit Sub
    MoYr = Split(cboCutOff.Text, "-")
    txtYr.Text = Trim(MoYr(1))
    txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & " 1, " & txtYr.Text))
    Call DisplaySales
  
End Sub
Private Sub chkSel_Click()
    
    If chkSel.Value = 1 Then
        flxSVRDay.Visible = True
        flxSPDay.Visible = False
        chkSel.Caption = "SERVICE CENTER"
        stabSales.TabCaption(2) = "SERVICE CENTER DETAILS"
        stabSales.Tab = 1
        
    Else
        flxSPDay.Visible = True
        flxSVRDay.Visible = False
        chkSel.Caption = "SPAREPARTS"
        stabSales.TabCaption(2) = "SPAREPARTS DETAILS"
        stabSales.Tab = 0
       
    End If
     Timer3.Enabled = True
End Sub

Private Sub cmdRep_Click()
    ReProc = True
    Call cboCutOff_Click
   
End Sub

Private Sub cmdTrans_Click()
     Unload Me
     MotoRep.tbMain.Visible = True
End Sub

'Private Sub cmdNext_Click()
'    cmdNext.Tag = "2"
'    cmdNext.Visible = False
'    cmdPrev.Left = cmdNext.Left
'    cmdPrev.Top = cmdNext.Top
'    cmdPrev.Visible = True
'    dTo.Visible = True
'    cmdRefresh.Visible = True
'    dTo.Value = dFrom.Value
'End Sub

'Private Sub cmdPrev_Click()
'    cmdNext.Tag = "1"
'    cmdNext.Visible = True
'    cmdPrev.Visible = False
'    dTo.Visible = False
'    cmdRefresh.Visible = False
    'Call RefreshRecord(1)
'End Sub
Private Sub flxSP_EnterCell()
    ActRow = flxSP.Row
End Sub
Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuINV)
    End If
End Sub
Private Sub Form_Load()
    Call LoadMos
    flxSPDay.Visible = True
    frmDay.Caption = "Sales For the Day " & Format(Date, "dddd mmmm dd, yyyy")
    stabSales.TabCaption(2) = "SPAREPARTS DETAILS"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsMSALESADJ.State = adStateOpen Then rsMSALESADJ.Close
End Sub
Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, j As Integer, r As Long
    
    'lblHeading(2).Left = Frame2.Left + (Frame2.Width - lblHeading(2).Width) - 30
    'lblHeading(3).Left = Frame2.Left + (Frame2.Width - lblHeading(3).Width)
    
    With flxHead
        .ColWidth(0) = 0
        
        For i = 1 To 14
            .ColWidth(i) = (.Width - .ColWidth(0)) / 14
            .ColAlignment(i) = 4
        Next i
        
        .TextMatrix(0, 1) = "DATE"
        .TextMatrix(1, 1) = "DATE"
        
        .TextMatrix(0, 2) = "C A S H"
        .TextMatrix(0, 3) = "C A S H"
        
        .TextMatrix(0, 4) = "O N - A C C O U N T"
        .TextMatrix(0, 5) = "O N - A C C O U N T"
        .TextMatrix(0, 6) = "O N - A C C O U N T"
        .TextMatrix(0, 7) = "O N - A C C O U N T"
        .TextMatrix(0, 8) = "O N - A C C O U N T"
        .TextMatrix(0, 9) = "O N - A C C O U N T"
        
        .TextMatrix(0, 10) = "SALES"
        .TextMatrix(1, 10) = "RETURN"
        
        .TextMatrix(0, 11) = "ADJUSTMENT"
        .TextMatrix(0, 12) = "ADJUSTMENT"
       
        .TextMatrix(1, 11) = "DEBIT"
        .TextMatrix(1, 12) = "CREDIT"
        
        .TextMatrix(1, 2) = "WALK-IN"
        .TextMatrix(1, 3) = "WHL. SALE"
        .TextMatrix(1, 4) = "RBLT-PLAN"
        .TextMatrix(1, 5) = "LGC-MKTG."
        .TextMatrix(1, 6) = "LGC-FIN"
        .TextMatrix(1, 7) = "WHL. SALE"
        .TextMatrix(1, 8) = "EMP.ACCT"
        .TextMatrix(1, 9) = "CORT-ACCT"
        
        .TextMatrix(0, 13) = "OTHERS"
        .TextMatrix(1, 13) = "OTHERS"
        
        .TextMatrix(0, 14) = "GRAND"
        .TextMatrix(1, 14) = "TOTAL"
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
        .MergeCol(9) = True
        .MergeCol(10) = True
        .MergeCol(11) = True
        .MergeCol(12) = True
        .MergeCol(13) = True
        .MergeCol(14) = True
       
    End With
    flxSP.Width = flxHead.Width
    With flxSP
        .Top = (flxHead.Top + flxHead.Height)
        .ColWidth(0) = 0
        For i = 1 To 14
            .ColWidth(i) = .Width / 14
            .ColAlignment(i) = 4
        Next i
    End With
    
    With flxSPDay
        .ColWidth(0) = 0
        
        For i = 1 To 13
            .ColWidth(i) = (.Width - .ColWidth(0)) / 13
            .ColAlignment(i) = 4
        Next i
        
        .TextMatrix(0, 1) = "C A S H"
        .TextMatrix(0, 2) = "C A S H"
        
        .TextMatrix(0, 3) = "O N - A C C O U N T"
        .TextMatrix(0, 4) = "O N - A C C O U N T"
        .TextMatrix(0, 5) = "O N - A C C O U N T"
        .TextMatrix(0, 6) = "O N - A C C O U N T"
        .TextMatrix(0, 7) = "O N - A C C O U N T"
        .TextMatrix(0, 9) = "O N - A C C O U N T"
        
        .TextMatrix(0, 9) = "SALES"
        .TextMatrix(1, 9) = "RETURN"
        
        .TextMatrix(0, 10) = "ADJUSTMENT"
        .TextMatrix(0, 11) = "ADJUSTMENT"
       
        .TextMatrix(1, 10) = "DEBIT"
        .TextMatrix(1, 11) = "CREDIT"
        
        .TextMatrix(1, 1) = "WALK-IN"
        .TextMatrix(1, 2) = "WHL. SALE"
        .TextMatrix(1, 3) = "RBLT-PLAN"
        .TextMatrix(1, 4) = "LGC-MKTG."
        .TextMatrix(1, 5) = "LGC-FIN"
        .TextMatrix(1, 6) = "WHL. SALE"
        .TextMatrix(1, 7) = "EMP.ACCT"
        .TextMatrix(1, 8) = "CORT-ACCT"
        
        .TextMatrix(0, 12) = "OTHERS"
        .TextMatrix(1, 12) = "OTHERS"
        
        .TextMatrix(0, 13) = "T O T A L"
        .TextMatrix(1, 13) = "T O T A L"
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
        .MergeCol(9) = True
        .MergeCol(10) = True
        .MergeCol(11) = True
        .MergeCol(12) = True
        .MergeCol(13) = True
        
       
    End With
    
    With flxHead2
        .ColWidth(0) = 0
        For i = 1 To 10
            .ColWidth(i) = (.Width - .ColWidth(0)) / 10
            .ColAlignment(i) = 4
        Next i
        .TextMatrix(0, 1) = "DATE"
        .TextMatrix(1, 1) = "DATE"
        
        .TextMatrix(0, 2) = "C A S H"
        .TextMatrix(1, 2) = "WALK-IN"
        
        .TextMatrix(0, 3) = "O N - A C C O U N T"
        .TextMatrix(0, 4) = "O N - A C C O U N T"
        .TextMatrix(0, 5) = "O N - A C C O U N T"
        .TextMatrix(0, 6) = "O N - A C C O U N T"
               
        .TextMatrix(0, 7) = "ADJUSTMENT"
        .TextMatrix(0, 8) = "ADJUSTMENT"
        .TextMatrix(1, 7) = "DEBIT"
        .TextMatrix(1, 8) = "CREDIT"
        
        .TextMatrix(1, 3) = "LGC- MKTG"
        .TextMatrix(1, 4) = "LGC - FIN"
        .TextMatrix(1, 5) = "EMP. ACCT"
        .TextMatrix(1, 6) = "CORT."
        
        .TextMatrix(0, 9) = "OTHERS"
        .TextMatrix(1, 9) = "OTHERS"
        
        .TextMatrix(0, 10) = "GRAND"
        .TextMatrix(1, 10) = "TOTAL"
       
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
        .MergeCol(9) = True
        .MergeCol(10) = True
    End With
    
    With flxSP2
        .Top = (flxHead2.Top + flxHead2.Height)
        .ColWidth(0) = 0
        For i = 1 To 10
            .ColWidth(i) = .Width / 10
            .ColAlignment(i) = 4
        Next i
    End With
    
    With flxSVRDay
        .ColWidth(0) = 0
        For i = 1 To 9
            .ColWidth(i) = (.Width - .ColWidth(0)) / 9
            .ColAlignment(i) = 4
        Next i
        
        .TextMatrix(0, 1) = "C A S H"
        .TextMatrix(1, 1) = "WALK-IN"
        
        .TextMatrix(0, 2) = "O N - A C C O U N T"
        .TextMatrix(0, 3) = "O N - A C C O U N T"
        .TextMatrix(0, 4) = "O N - A C C O U N T"
        .TextMatrix(0, 5) = "O N - A C C O U N T"
               
        .TextMatrix(0, 6) = "ADJUSTMENT"
        .TextMatrix(0, 7) = "ADJUSTMENT"
        .TextMatrix(1, 6) = "DEBIT"
        .TextMatrix(1, 7) = "CREDIT"
        
        .TextMatrix(1, 2) = "LGC- MKTG"
        .TextMatrix(1, 3) = "LGC - FIN"
        .TextMatrix(1, 4) = "EMP. ACCT"
        .TextMatrix(1, 5) = "CORT."
        
        .TextMatrix(0, 8) = "OTHERS"
        .TextMatrix(1, 8) = "OTHERS"
        
        .TextMatrix(0, 9) = "T O T A L"
        .TextMatrix(1, 9) = "T O T A L"
       
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
        .MergeCol(9) = True
        
    End With
    
End Sub
Private Sub stabSales_Click(PreviousTab As Integer)
    If stabSales.Tab = 1 Then
        chkSel.Value = 1
    ElseIf stabSales.Tab = 0 Then
        chkSel.Value = 0
    End If
End Sub
Private Sub Timer1_Timer()
    'cboCutOff.ListIndex = 0
    
    If chkSel.Value = 0 Then
        cboCutOff.Text = txtsp.Text
    Else
        cboCutOff.Text = txtsp2.Text
    End If
    
    Timer1.Enabled = False
    
End Sub
Private Sub DisplaySales()
    Dim Qry As String
    Dim dStart As Date, dEnd As Date
    Dim TtlSP As Currency, TtlSrv As Currency, GTtl As Currency
    
    dStart = DateSerial(txtYr.Text, txtMo.Text, 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    
    If chkSel.Value = 0 Then
        If ReProc = True Then
            LYRcon.Execute "DELETE FROM msalesbrkdownsp Where (((msalesbrkdownsp.SDATE) >= '" & Format(dStart, "yyyy-mm-dd") & "' And (msalesbrkdownsp.SDATE) <= '" & Format(dEnd, "yyyy-mm-dd") & "') And ((msalesbrkdownsp.SPTOTAL) > 0)) ORDER BY msalesbrkdownsp.SDATE;"
            DoEvents
        End If
            Qry = "SELECT msalesbrkdownsp.* From msalesbrkdownsp Where msalesbrkdownsp.SDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' And msalesbrkdownsp.SDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' And msalesbrkdownsp.SPTOTAL > 0 ORDER BY msalesbrkdownsp.SDATE;"
            txtSQL.Text = "SELECT msalesbrkdownsp.SDATE, msalesbrkdownsp.WHLSALE, msalesbrkdownsp.RBLTPLAN, msalesbrkdownsp.LGCMKTG, msalesbrkdownsp.LGCFIN, msalesbrkdownsp.WHLACCT, msalesbrkdownsp.EMPACCT, msalesbrkdownsp.CORTACCT, msalesbrkdownsp.SRETURN, msalesbrkdownsp.AdjDebit, " & _
                      "msalesbrkdownsp.AdjCredit, msalesbrkdownsp.SOTHERS, msalesbrkdownsp.SPTotal From msalesbrkdownsp " & _
                      "Where (((msalesbrkdownsp.SDATE) >= '" & Format(dStart, "yyyy-mm-dd") & "' And (msalesbrkdownsp.SDATE) <= '" & Format(dEnd, "yyyy-mm-dd") & "') And ((msalesbrkdownsp.SPTOTAL) > 0)) ORDER BY msalesbrkdownsp.SDATE;"
    Else
        If ReProc = True Then
            LYRcon.Execute "DELETE FROM msalesbrkdownsvr Where (((msalesbrkdownsvr.SDATE) >= '" & Format(dStart, "yyyy-mm-dd") & "' And (msalesbrkdownsvr.SDATE) <= '" & Format(dEnd, "yyyy-mm-dd") & "') And ((msalesbrkdownsvr.svrTOTAL) > 0)) ORDER BY msalesbrkdownsvr.SDATE;"
            DoEvents
        End If
            Qry = "SELECT msalesbrkdownsvr.* From msalesbrkdownsvr Where msalesbrkdownsvr.SDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' And msalesbrkdownsvr.SDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' And msalesbrkdownsvr.SvrTOTAL > 0 ORDER BY msalesbrkdownsvr.SDATE;"
            txtSQL.Text = "SELECT msalesbrkdownsvr.SDATE, msalesbrkdownsvr.WALKIN, msalesbrkdownsvr.LGCMKTG, msalesbrkdownsvr.LGCFIN, msalesbrkdownsvr.EMPACCT, msalesbrkdownsvr.CORTACCT, msalesbrkdownsvr.ADJDEBIT, msalesbrkdownsvr.ADJCREDIT, msalesbrkdownsvr.SOTHERS, " & _
            "msalesbrkdownsvr.SVRTOTAL From msalesbrkdownsvr " & _
            "Where (((msalesbrkdownsvr.SDATE) >= '" & Format(dStart, "yyyy-mm-dd") & "' And (msalesbrkdownsvr.SDATE) <= '" & Format(dEnd, "yyyy-mm-dd") & "') And ((msalesbrkdownsvr.SVRTOTAL) > 0)) ORDER BY msalesbrkdownsvr.SDATE;"
    End If
    
        If rsMSALESADJ.State = adStateOpen Then rsMSALESADJ.Close
        rsMSALESADJ.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsMSALESADJ.RecordCount = 0 Then
            
            If chkSel.Value = 0 Then
                flxSP.Rows = 0
            Else
                flxSP2.Rows = 0
            End If
            
            Ans = MsgBox("Do you want to Process Sales for then Period " & cboCutOff.Text & "?", vbQuestion + vbYesNo, "Sales")
            If Ans = vbNo Then Exit Sub
            Call ProcessSales
            Exit Sub
        End If
   
    With rsMSALESADJ
        .MoveFirst
       
        
        If chkSel.Value = 0 Then
            flxSP.Rows = 0
            txtsp.Text = ""
            txtsp.Text = cboCutOff.Text
            Do While Not .EOF
                'SDATE        WALKIN    WHLSALE    RBLTPLAN   LGCMKTG LGCFIN   WHLACCT    EMPACCT  CORTACCT SRETURN   AdjDebit   AdjCredit  SOTHERS  SPTotal
                flxSP.AddItem vbTab & Format(!SDATE, "mm/dd/yyyy") & vbTab & Format(!WALKIN, "#,##0.00") & vbTab & Format(!WHLSALE, "#,##0.00") & vbTab & Format(!RBLTPLAN, "#,##0.00") & vbTab & Format(!LGCMKTG, "#,##0.00") & vbTab & Format(!LGCFIN, "#,##0.00") & vbTab & Format(!WHLACCT, "#,##0.00") & vbTab & Format(!EMPACCT, "#,##0.00") & vbTab & Format(!CORTACCT, "#,##0.00") & vbTab & Format(!SRETURN, "#,##0.00") & vbTab & Format(!AdjDebit, "#,##0.00") & vbTab & Format(!AdjCredit, "#,##0.00") & vbTab & Format(!SOTHERS, "#,##0.00") & vbTab & Format(!SPTotal, "#,##0.00")
                .MoveNext
            Loop
        Else
            flxSP2.Rows = 0
            txtsp2.Text = ""
            txtsp2.Text = cboCutOff.Text
            Do While Not .EOF
                '                             Date                                   SPWwcash                             SrvLGCAcct                              SrvFinAcct                            DownPay                               SrvCorpAcct                                SpDbt                                SvcCdt                                  Others                                  Total
                flxSP2.AddItem vbTab & Format(!SDATE, "mm/dd/yyyy") & vbTab & Format(!WALKIN, "#,##0.00") & vbTab & Format(!LGCMKTG, "#,##0.00") & vbTab & Format(!LGCFIN, "#,##0.00") & vbTab & Format(!EMPACCT, "#,##0.00") & vbTab & Format(!CORTACCT, "#,##0.00") & vbTab & Format(!AdjDebit, "#,##0.00") & vbTab & Format(!AdjCredit, "#,##0.00") & vbTab & Format(!SOTHERS, "#,##0.00") & vbTab & Format(!SVRTOTAL, "#,##0.00")
               .MoveNext
            Loop
        End If
        
       .Close
    End With
    ActRow = 1
   If chkSel.Value = 0 Then
        Qry = "SELECT Sum(msalesbrkdownsp.SPTOTAL) AS SumOfTTLSALESP From msalesbrkdownsp WHERE msalesbrkdownsp.SDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And msalesbrkdownsp.SDATE<='" & Format(dEnd, "yyyy-mm-dd") & "';"
   Else
        Qry = "SELECT Sum(msalesbrkdownsvr.SvrTOTAL) AS SumOfTTLSALESVR From msalesbrkdownsvr WHERE msalesbrkdownsvr.SDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And msalesbrkdownsvr.SDATE<='" & Format(dEnd, "yyyy-mm-dd") & "';"
   End If
   lblMoColl.Caption = "Total Sales for the Month: " & Format(Val(GetFieldValue(Qry)), "#,##0.00")

End Sub
Private Sub ProcessSales()
    Dim Qry As String, i As Integer, Ans As Integer, pPct As Single, SvcDisc As Currency, SvcAcct As Currency
    Dim dStart As Date, dEnd As Date, TmpD As Date
    
    
    dStart = DateSerial(txtYr.Text, txtMo.Text, 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
    
    flxSPPro.Rows = 1
    flxSPSrv.Rows = 1
    TmpD = dStart
    Do While TmpD <= dEnd
        If chkSel.Value = 0 Then
            flxSPPro.AddItem TmpD
        Else
            flxSPSrv.AddItem TmpD
        End If
        TmpD = DateAdd("d", 1, TmpD)
    Loop
    
    'Lock objects
    Frame1.Enabled = False
    If chkSel.Value = 0 Then
        flxSPPro.Enabled = False
        flxSPSrv.Enabled = True
    Else
        flxSPPro.Enabled = True
        flxSPSrv.Enabled = False
    End If
    Pbar.Value = 0
    If chkSel.Value = 0 Then
        Pbar.Max = (flxSPPro.Rows - 1) * 2
    Else
        Pbar.Max = (flxSPSrv.Rows - 1) * 2
    End If
    Pbar.Visible = True
    frmwait.Visible = True
    'Timer3.Enabled = True
    With flxSPPro
        If chkSel.Value = 0 Then
             stabSales.TabCaption(2) = "SPAREPARTS DETAILS"
             For i = 1 To .Rows - 1
                  'Monthly
                  'SP Cash
                  Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP_CASH From payments INNER JOIN spinvoice ON (payments.INVOICENO = spinvoice.SPINVNO) " & _
                      " Where payments.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL AND payments.CUSTOMERNAME = 'CASH' AND spinvoice.CTYPE <> 'WHOLESALE';"
                  SPWcash = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 1) = Format(SPWcash, "#,##0.00")
                  
                  'SP Cash Wholesale
                  Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP_CW From payments INNER JOIN spinvoice ON (payments.INVOICENO = spinvoice.SPINVNO) " & _
                      " Where payments.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL AND payments.CUSTOMERNAME = 'CASH' AND spinvoice.CTYPE = 'WHOLESALE';"
                  SPWwcash = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 2) = Format(SPWwcash, "#,##0.00")
                  
                  'SP On Acct
                  '67=CUSTOMERID 2h rebuilt Plant
                  Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_2H From SPINVOICE " & _
                        "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CUSTOMERID=67 AND SPINVOICE.INVDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                  SP2HAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 3) = Format(SP2HAcct, "#,##0.00")
                  'SP On Acct
                  'LGC Marketing not include rebuilt Plant
                  Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_LGC From SPINVOICE " & _
                        "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CUSTOMERID<>67 AND SPINVOICE.CTYPE='LGC MARKETING' AND SPINVOICE.INVDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                  SPLGCAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 4) = Format(SPLGCAcct, "#,##0.00")
                  
                  'SP On Acct
                  'LGC FINANCING ON ACCOUNT
                  Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_LGCF From SPINVOICE " & _
                        "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CTYPE='LGC FINANCING' AND SPINVOICE.INVDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                  SPLGCFAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 5) = Format(SPLGCFAcct, "#,##0.00")
            
                  'SP On Acct
                  'WHOLESALE ON ACCOUNT
                  Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_WA From SPINVOICE " & _
                        "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CTYPE='WHOLESALE' AND SPINVOICE.INVDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                  SPLGCWAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 6) = Format(SPLGCWAcct, "#,##0.00")
                  
                  'SP On Acct
                  'INDIVIDUAL ON ACCOUNT
                  Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_AI From SPINVOICE " & _
                        "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CTYPE='INDIVIDUAL' AND SPINVOICE.INVDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                  SPLGCIAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 7) = Format(SPLGCIAcct, "#,##0.00")
                  
                  'SP On Acct
                  'CORPORATE ON ACCOUNT
                  Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_CA From SPINVOICE " & _
                        "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CTYPE='CORPORATE' AND SPINVOICE.INVDATE='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                  SPLGCIAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 8) = Format(SPLGCIAcct, "#,##0.00")
                  
                  'SP Sales Ret
                  Qry = "SELECT Sum(SRETURNSP.TOTALAMT) As SumOfTOTALAMT From SRETURNSP LEFT OUTER JOIN spinvoice ON (SRETURNSP.INVOICE = spinvoice.SPINVNO) " & _
                        "Where spinvoice.INVDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                  SRet = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 9) = Format(SRet, "#,##0.00")
                  
                  'SP Insert Debit amount
                  Qry = "SELECT SUM(adjdmcm.DEBIT) As SumOfTtlDebit From adjdmcm LEFT OUTER JOIN spinvoice ON (adjdmcm.REFNO = spinvoice.SPINVNO) " & _
                        "Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND adjdmcm.DEBIT > 0 AND adjdmcm.REASON <> 'CANCELLED' " & _
                        "AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
                  SpDbt = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 10) = Format(SpDbt, "#,##0.00")
                  
                  'Insert Credit amount
                  Qry = "SELECT SUM(adjdmcm.CREDIT) As SumOfTtlCredit From adjdmcm LEFT OUTER JOIN spinvoice ON (adjdmcm.REFNO = spinvoice.SPINVNO) " & _
                        "Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND adjdmcm.CREDIT > 0 AND adjdmcm.REASON <> 'CANCELLED' " & _
                        "AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
                  SpCdt = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 11) = Format(SpCdt, "#,##0.00")
                   
                  Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS_SP From PAYMENTS " & _
                       "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') AND (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
            
                  SPOthers = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 12) = Format(SPOthers, "#,##0.00")
                  
                  'SP Total
                  .TextMatrix(i, 13) = Format((((SPWcash + SPWwcash + SP2HAcct + SPLGCAcct + SPLGCFAcct + SPLGCWAcct + SPLGCIAcct + SPLGCIAcct + Others) - SRet) + SpDbt) - SpCdt, "#,##0.00")
                    
                   Pbar.Value = Pbar.Value + 1
                   DoEvents
                   pPct = (Pbar.Value / Pbar.Max)
                   lblTotal.Caption = Format$(pPct, "##0.0%")
            Next i
        End If
    End With
    
    With flxSPSrv
        If chkSel.Value = 1 Then
            stabSales.TabCaption(2) = "SERVICE CENTER DETAILS"
            For i = 1 To .Rows - 1
                  'Service Cash
                  Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE_WC From PAYMENTS " & _
                        "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') And ((PAYMENTS.CUSTOMERID) = 0)) HAVING (((Sum(PAYMENTS.SERVICE))>0));"
                  SrvCash = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 1) = Format(SrvCash, "#,##0.00")
                  
                  'ON ACCOUNT LGC MARKETING
                  Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) As SumOfTOTALCOST From SRVCJOBORDER INNER JOIN customer ON (SRVCJOBORDER.CUSTOMERID = customer.CUSTOMERID) " & _
                        "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND  customer.CTYPE = 'LGC MARKETING';"
                  SrvLGCAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 2) = Format(SrvLGCAcct, "#,##0.00")
                  
                  'ON ACCOUNT LGC FINANCING
                  Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) As SumOfTOTALCOST From SRVCJOBORDER INNER JOIN customer ON (SRVCJOBORDER.CUSTOMERID = customer.CUSTOMERID) " & _
                        "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND Customer.CTYPE = 'LGC FINANCING';"
                  SrvFinAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 3) = Format(SrvFinAcct, "#,##0.00")
                   
                  'ON ACCOUNT EMPLOYEE ACCOUNT
                  Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) As SumOfTOTALCOST From SRVCJOBORDER INNER JOIN customer ON (SRVCJOBORDER.CUSTOMERID = customer.CUSTOMERID) " & _
                        "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND Customer.CTYPE = 'INDIVIDUAL';"
                  
                  SrvEmpAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 4) = Format(SrvEmpAcct, "#,##0.00")
                  
                  'CORPORATE ACCOUNT
                  Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) As SumOfTOTALCOST From SRVCJOBORDER INNER JOIN customer ON (SRVCJOBORDER.CUSTOMERID = customer.CUSTOMERID) " & _
                        "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND Customer.CTYPE = 'CORPORATE';"
                  
                  SrvCorpAcct = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 5) = Format(SrvCorpAcct, "#,##0.00")
                  
                  'Service Insert Debit amount
                  Qry = "SELECT SUM(adjdmcm.DEBIT) As SvcSumTtlDebit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
                        "Where payments.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' AND " & _
                        "adjdmcm.TRANSTYPE = 'JOB ORDER' AND adjdmcm.REASON <> 'CANCELLED'"
                  SvcDbt = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 6) = Format(SvcDbt, "#,##0.00")
                 
                  'Service Insert Credit amount
                  Qry = "SELECT SUM(adjdmcm.CREDIT) As SvcSumTtlCredit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
                      "Where payments.PAYDATE = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' AND " & _
                      "adjdmcm.TRANSTYPE = 'JOB ORDER' AND adjdmcm.REASON <> 'CANCELLED'"
                  SvcCdt = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 7) = Format(SvcCdt, "#,##0.00")
                  'Other Downpayment
                  Qry = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDownPayment From PAYMENTS " & _
                       "Where (((PAYMENTS.PAYDATE) = '" & Format(.TextMatrix(i, 0), "yyyy-mm-dd") & "') AND (PAYMENTS.DOWNPAYMENT > 0) AND (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.DOWNPAYMENT))<>0));"
            
                  DownPay = Val(GetFieldValue(Qry))
                  .TextMatrix(i, 8) = Format(DownPay, "#,##0.00")
        
                  .TextMatrix(i, 9) = Format((((SPWwcash + SrvLGCAcct + SrvFinAcct + DownPay + SrvCorpAcct) + SpDbt) - SvcCdt), "#,##0.00")
            
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                    pPct = (Pbar.Value / Pbar.Max)
                    lblTotal.Caption = Format$(pPct, "##0.0%")
            Next i
        End If
    End With
    
    If chkSel = 0 Then
        For i = 1 To flxSPPro.Rows - 1
            If Val(flxSPPro.TextMatrix(i, 13)) > 0 Then
                    Qry = "SELECT msalesbrkdownsp.* From msalesbrkdownsp WHERE msalesbrkdownsp.SDATE='" & Format(flxSPPro.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                    
                    If CtrlNumExist(Qry) Then
                        LYRcon.Execute "UPDATE msalesbrkdownsp SET WALKIN = " & Val(Replace(flxSPPro.TextMatrix(i, 1), ",", "")) & ",  WHLSALE = " & Val(Replace(flxSPPro.TextMatrix(i, 2), ",", "")) & ", RBLTPLAN = " & Val(Replace(flxSPPro.TextMatrix(i, 3), ",", "")) & ", LGCMKTG = " & Val(Replace(flxSPPro.TextMatrix(i, 4), ",", "")) & ", LGCFIN = " & Val(Replace(flxSPPro.TextMatrix(i, 5), ",", "")) & ", WHLACCT = " & Val(Replace(flxSPPro.TextMatrix(i, 6), ",", "")) & ", EMPACCT = " & Val(Replace(flxSPPro.TextMatrix(i, 7), ",", "")) & ", SRETURN = " & Val(Replace(flxSPPro.TextMatrix(i, 8), ",", "")) & ", CORTACCT = " & Val(Replace(flxSPPro.TextMatrix(i, 9), ",", "")) & ", ADJDEBIT = " & Val(Replace(flxSPPro.TextMatrix(i, 10), ",", "")) & ", ADJCREDIT = " & Val(Replace(flxSPPro.TextMatrix(i, 11), ",", "")) & ", SOTHERS = " & Val(Replace(flxSPPro.TextMatrix(i, 12), ",", "")) & ", SPTOTAL = " & Val(Replace(flxSPPro.TextMatrix(i, 13), ",", "")) & _
                            " WHERE msalesbrkdownsp.SDATE='" & Format(flxSPPro.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                    Else
                         LYRcon.Execute "INSERT INTO msalesbrkdownsp(SDATE, WALKIN, WHLSALE, RBLTPLAN, LGCMKTG, LGCFIN, WHLACCT, EMPACCT, CORTACCT, SRETURN, ADJDEBIT, ADJCREDIT, SOTHERS, SPTOTAL)VALUES('" & Format(flxSPPro.TextMatrix(i, 0), "yyyy-mm-dd") & "', " & Val(Replace(flxSPPro.TextMatrix(i, 1), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 2), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 3), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 4), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 5), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 6), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 7), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 8), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 9), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 10), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 11), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 12), ",", "")) & ", " & Val(Replace(flxSPPro.TextMatrix(i, 13), ",", "")) & ")"
                    End If
                    
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                    pPct = (Pbar.Value / Pbar.Max)
                    lblTotal.Caption = Format$(pPct, "##0.0%")
            End If
        Next i
    Else
        For i = 1 To flxSPSrv.Rows - 1
            If Val(flxSPSrv.TextMatrix(i, 9)) > 0 Then
                    Qry = "SELECT msalesbrkdownsvr.* From msalesbrkdownsvr WHERE msalesbrkdownsvr.SDATE='" & Format(flxSPSrv.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                    
                    If CtrlNumExist(Qry) Then
                        LYRcon.Execute "UPDATE msalesbrkdownsvr SET WALKIN = " & Val(Replace(flxSPSrv.TextMatrix(i, 1), ",", "")) & ", LGCMKTG = " & Val(Replace(flxSPSrv.TextMatrix(i, 2), ",", "")) & ", LGCFIN = " & Val(Replace(flxSPSrv.TextMatrix(i, 3), ",", "")) & ", EMPACCT = " & Val(Replace(flxSPPro.TextMatrix(i, 4), ",", "")) & ", CORTACCT = " & Val(Replace(flxSPSrv.TextMatrix(i, 5), ",", "")) & ", ADJDEBIT = " & Val(Replace(flxSPSrv.TextMatrix(i, 6), ",", "")) & ", ADJCREDIT = " & Val(Replace(flxSPSrv.TextMatrix(i, 7), ",", "")) & ", SOTHERS = " & Val(Replace(flxSPSrv.TextMatrix(i, 8), ",", "")) & ", SVRTOTAL = " & Val(Replace(flxSPSrv.TextMatrix(i, 9), ",", "")) & _
                                  " WHERE msalesbrkdownsvr.SDATE='" & Format(flxSPSrv.TextMatrix(i, 0), "yyyy-mm-dd") & "';"
                    Else
                        LYRcon.Execute "INSERT INTO msalesbrkdownsvr(SDATE, WALKIN, LGCMKTG, LGCFIN, EMPACCT, CORTACCT, ADJDEBIT, ADJCREDIT, SOTHERS, SVRTOTAL)VALUES('" & Format(flxSPSrv.TextMatrix(i, 0), "yyyy-mm-dd") & "', " & Val(Replace(flxSPSrv.TextMatrix(i, 1), ",", "")) & ", " & Val(Replace(flxSPSrv.TextMatrix(i, 2), ",", "")) & ", " & Val(Replace(flxSPSrv.TextMatrix(i, 3), ",", "")) & ", " & Val(Replace(flxSPSrv.TextMatrix(i, 4), ",", "")) & ", " & Val(Replace(flxSPSrv.TextMatrix(i, 5), ",", "")) & ", " & Val(Replace(flxSPSrv.TextMatrix(i, 6), ",", "")) & ", " & Val(Replace(flxSPSrv.TextMatrix(i, 7), ",", "")) & ", " & Val(Replace(flxSPSrv.TextMatrix(i, 8), ",", "")) & ", " & Val(Replace(flxSPSrv.TextMatrix(i, 9), ",", "")) & ")"
                    End If
                    
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                    pPct = (Pbar.Value / Pbar.Max)
                    lblTotal.Caption = Format$(pPct, "##0.0%")
            End If
        Next i
    End If
    
    DoEvents
    
    MsgBox "Finish processing Sales for " & MonthName(txtMo.Text) & " " & txtYr.Text & "!", vbInformation + vbOKOnly, "Update Complete"
    Pbar.Visible = False
    lblTotal.Caption = ""
    Frame1.Enabled = True
    flxSP.Enabled = True
    frmwait.Visible = False
    'Timer3.Enabled = False
    Call DisplaySales
End Sub
Private Sub LoadMos()
    Dim Mo As Integer, Yr As Integer, MoYr As Boolean, TmpD As Date
 
    Mo = 1
    Yr = 2008
    'cboCutOff.AddItem "==Exclusive Date=="
    Do
        cboCutOff.AddItem MonthName(Mo) & "-" & Yr, 0
        TmpD = DateAdd("m", 1, DateSerial(Yr, Mo, 1))
        Mo = Month(TmpD)
        Yr = Year(TmpD)
        MoYr = IIf(TmpD <= DateSerial(Year(Date), Month(Date), 1), True, False)
    Loop While MoYr
        
End Sub
Private Sub LoadDaySales()
    
    Dim Qry As String, dStart As Date, dEnd As Date, SvcCdt As Single
    Dim SPWwcash As Single, SrvLGCAcct As Single, SrvFinAcct As Single, DownPay As Single, SrvCorpAcct As Single
    Dim SPWcash As Single, SP2HAcct As Single, SPLGCAcct As Single, SPLGCFAcct As Single, SPLGCWAcct As Single
    Dim SPLGCIAcct As Single, Others As Single, SRet As Single, SpDbt As Single, SpCdt As Single, SvcDbt As Single
    

    'Query if the record exist the table
    
    If chkSel.Value = 0 Then
        Qry = "SELECT msalesbrkdownsp.SPTotal From msalesbrkdownsp Where msalesbrkdownsp.SPTOTAL > 0 AND msalesbrkdownsp.SDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
    Else
        Qry = "SELECT msalesbrkdownsvr.SvrTotal From msalesbrkdownsvr Where msalesbrkdownsvr.SvrTOTAL > 0 AND msalesbrkdownsvr.SDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
    End If
    If CtrlNumExist(Qry) Then
        Call LoadTodaySales
    Else
        If chkSel.Value = 0 Then
              'Monthly
              stabSales.TabCaption(2) = "SPAREPARTS DETAILS"
              'dStart = DateSerial(Year(Now), Month(Now), 1)
              'dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
              
              'SP Cash
              Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP_CASH From payments INNER JOIN spinvoice ON (payments.INVOICENO = spinvoice.SPINVNO) " & _
                  " Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL AND payments.CUSTOMERNAME = 'CASH' AND spinvoice.CTYPE <> 'WHOLESALE';"
              SPWcash = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 1) = Format(SPWcash, "#,##0.00")
              
              'SP Cash Wholesale
              Qry = "SELECT SUM(payments.SPAREPARTS) As SumOfSP_CW From payments INNER JOIN spinvoice ON (payments.INVOICENO = spinvoice.SPINVNO) " & _
                  " Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.INVOICENO IS NOT NULL AND payments.CUSTOMERNAME = 'CASH' AND spinvoice.CTYPE = 'WHOLESALE';"
              SPWwcash = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 2) = Format(SPWwcash, "#,##0.00")
              
              'SP On Acct
              '67=CUSTOMERID 2h rebuilt Plant
              Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_2H From SPINVOICE " & _
                    "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CUSTOMERID=67 AND SPINVOICE.INVDATE='" & Format(Date, "yyyy-mm-dd") & "';"
              SP2HAcct = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 3) = Format(SP2HAcct, "#,##0.00")
              'SP On Acct
              'LGC Marketing not include rebuilt Plant
              Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_LGC From SPINVOICE " & _
                    "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CUSTOMERID<>67 AND SPINVOICE.CTYPE='LGC MARKETING' AND SPINVOICE.INVDATE='" & Format(Date, "yyyy-mm-dd") & "';"
              SPLGCAcct = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 4) = Format(SPLGCAcct, "#,##0.00")
              
              'SP On Acct
              'LGC FINANCING STATEMENT
              Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_LGCF From SPINVOICE " & _
                    "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CTYPE='LGC FINANCING' AND SPINVOICE.INVDATE='" & Format(Date, "yyyy-mm-dd") & "';"
              SPLGCFAcct = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 5) = Format(SPLGCFAcct, "#,##0.00")
        
              'SP On Acct
              'WHOLESALE ON ACCOUNT
              Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_WA From SPINVOICE " & _
                    "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CTYPE='WHOLESALE' AND SPINVOICE.INVDATE='" & Format(Date, "yyyy-mm-dd") & "';"
              SPLGCWAcct = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 6) = Format(SPLGCWAcct, "#,##0.00")
              
              'SP On Acct
              'INDIVIDUAL ON ACCOUNT
              Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_AI From SPINVOICE " & _
                    "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CTYPE='INDIVIDUAL' AND SPINVOICE.INVDATE='" & Format(Date, "yyyy-mm-dd") & "';"
              SPLGCIAcct = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 7) = Format(SPLGCIAcct, "#,##0.00")
              
              'SP On Acct
              'CORPORATE ON ACCOUNT
              Qry = "SELECT Sum(SPINVOICE.AMOUNT) AS SumOfAMOUNT_CA From SPINVOICE " & _
                    "WHERE SPINVOICE.INVTYPE='CREDIT' AND SPINVOICE.CTYPE='CORPORATE' AND SPINVOICE.INVDATE='" & Format(Date, "yyyy-mm-dd") & "';"
              SPLGCIAcct = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 8) = Format(SPLGCIAcct, "#,##0.00")
              
              'SP Sales Ret
              Qry = "SELECT Sum(SRETURNSP.TOTALAMT) As SumOfTOTALAMT From SRETURNSP LEFT OUTER JOIN spinvoice ON (SRETURNSP.INVOICE = spinvoice.SPINVNO) " & _
                    "Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
              SRet = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 9) = Format(SRet, "#,##0.00")
              
              'SP Insert Debit amount
              Qry = "SELECT SUM(adjdmcm.DEBIT) As SumOfTtlDebit From adjdmcm LEFT OUTER JOIN spinvoice ON (adjdmcm.REFNO = spinvoice.SPINVNO) " & _
                    "Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND adjdmcm.DEBIT > 0 AND adjdmcm.REASON <> 'CANCELLED' " & _
                    "AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
              SpDbt = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 10) = Format(SpDbt, "#,##0.00")
              
              'Insert Credit amount
              Qry = "SELECT SUM(adjdmcm.CREDIT) As SumOfTtlCredit From adjdmcm LEFT OUTER JOIN spinvoice ON (adjdmcm.REFNO = spinvoice.SPINVNO) " & _
                    "Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND adjdmcm.CREDIT > 0 AND adjdmcm.REASON <> 'CANCELLED' " & _
                    "AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
              SpCdt = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 11) = Format(SpCdt, "#,##0.00")
              'Others
              Qry = "SELECT Sum(PAYMENTS.OTHERS) AS SumOfOTHERS_SP From PAYMENTS " & _
                   "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') AND (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.OTHERS))<>0));"
        
              SPOthers = Val(GetFieldValue(Qry))
              flxSPDay.TextMatrix(2, 12) = Format(SPOthers, "#,##0.00")
              
              'SP Total
              flxSPDay.TextMatrix(2, 13) = Format((((SPWcash + SPWwcash + SP2HAcct + SPLGCAcct + SPLGCFAcct + SPLGCWAcct + SPLGCIAcct + SPLGCIAcct + Others) - SRet) + SpDbt) - SpCdt, "#,##0.00")
        Else
              stabSales.TabCaption(2) = "SERVICE CENTER DETAILS"
              'Service Cash
              Qry = "SELECT Sum(PAYMENTS.SERVICE) AS SumOfSERVICE_WC From PAYMENTS " & _
                    "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') And ((PAYMENTS.CUSTOMERID) = 0)) HAVING (((Sum(PAYMENTS.SERVICE))>0));"
              SrvCash = Val(GetFieldValue(Qry))
              flxSVRDay.TextMatrix(2, 1) = Format(SrvCash, "#,##0.00")
              
              'ON ACCOUNT LGC MARKETING
              Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) As SumOfTOTALCOST From SRVCJOBORDER INNER JOIN customer ON (SRVCJOBORDER.CUSTOMERID = customer.CUSTOMERID) " & _
                    "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(Date, "yyyy-mm-dd") & "' AND  customer.CTYPE = 'LGC MARKETING';"
              SrvLGCAcct = Val(GetFieldValue(Qry))
              flxSVRDay.TextMatrix(2, 2) = Format(SrvLGCAcct, "#,##0.00")
              
              'ON ACCOUNT LGC FINANCING
              Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) As SumOfTOTALCOST From SRVCJOBORDER INNER JOIN customer ON (SRVCJOBORDER.CUSTOMERID = customer.CUSTOMERID) " & _
                    "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(Date, "yyyy-mm-dd") & "' AND Customer.CTYPE = 'LGC FINANCING';"
              SrvFinAcct = Val(GetFieldValue(Qry))
              flxSVRDay.TextMatrix(2, 3) = Format(SrvFinAcct, "#,##0.00")
               
              'ON ACCOUNT EMPLOYEE ACCOUNT
              Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) As SumOfTOTALCOST From SRVCJOBORDER INNER JOIN customer ON (SRVCJOBORDER.CUSTOMERID = customer.CUSTOMERID) " & _
                    "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(Date, "yyyy-mm-dd") & "' AND Customer.CTYPE = 'INDIVIDUAL';"
              
              SrvEmpAcct = Val(GetFieldValue(Qry))
              flxSVRDay.TextMatrix(2, 4) = Format(SrvEmpAcct, "#,##0.00")
              
              'CORPORATE ACCOUNT
              Qry = "SELECT Sum(SRVCJOBORDER.TOTALCOST) As SumOfTOTALCOST From SRVCJOBORDER INNER JOIN customer ON (SRVCJOBORDER.CUSTOMERID = customer.CUSTOMERID) " & _
                    "WHERE SRVCJOBORDER.PAYTYPE='ON ACCOUNT' AND SRVCJOBORDER.DATEPROMISED='" & Format(Date, "yyyy-mm-dd") & "' AND Customer.CTYPE = 'CORPORATE';"
              
              SrvCorpAcct = Val(GetFieldValue(Qry))
              flxSVRDay.TextMatrix(2, 5) = Format(SrvCorpAcct, "#,##0.00")
              
              'Service Insert Debit amount
              Qry = "SELECT SUM(adjdmcm.DEBIT) As SvcSumTtlDebit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
                    "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' AND " & _
                    "adjdmcm.TRANSTYPE = 'JOB ORDER' AND adjdmcm.REASON <> 'CANCELLED'"
              SvcDbt = Val(GetFieldValue(Qry))
              flxSVRDay.TextMatrix(2, 6) = Format(SvcDbt, "#,##0.00")
             
              'Service Insert Credit amount
              Qry = "SELECT SUM(adjdmcm.CREDIT) As SvcSumTtlCredit From adjdmcm LEFT OUTER JOIN payments ON (adjdmcm.REFNO = payments.INVOICENO) " & _
                  "Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.SERVICE > 0 AND adjdmcm.ADJDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND adjdmcm.ADJDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "' AND " & _
                  "adjdmcm.TRANSTYPE = 'JOB ORDER' AND adjdmcm.REASON <> 'CANCELLED'"
              SvcCdt = Val(GetFieldValue(Qry))
              flxSVRDay.TextMatrix(2, 7) = Format(SvcCdt, "#,##0.00")
              
              'Other Downpayment
              Qry = "SELECT Sum(PAYMENTS.DOWNPAYMENT) AS SumOfDownPayment From PAYMENTS " & _
                   "Where (((PAYMENTS.PAYDATE) = '" & Format(Date, "yyyy-mm-dd") & "') AND (PAYMENTS.DOWNPAYMENT > 0) AND (Not (PAYMENTS.ORNO) Is Null)) HAVING (((Sum(PAYMENTS.DOWNPAYMENT))<>0));"
        
              DownPay = Val(GetFieldValue(Qry))
              flxSVRDay.TextMatrix(2, 8) = Format(DownPay, "#,##0.00")
    
              flxSVRDay.TextMatrix(2, 9) = Format((((SPWwcash + SrvLGCAcct + SrvFinAcct + DownPay + SrvCorpAcct) + SpDbt) - SvcCdt), "#,##0.00")
        End If
    End If
    Timer1.Enabled = True
    MotoRep.sBar.Panels(6).Text = "Sales Distribution"
    MotoRep.tbMain.Visible = False
    ActRow = 1
    
End Sub
Private Sub LoadTodaySales()
    Dim rsTSales As New ADODB.Recordset
    Dim Qry As String
    'Dim TtlSP As Currency, TtlSvc As Currency
 
    If chkSel.Value = 0 Then
        Qry = "SELECT msalesbrkdownsp.* From msalesbrkdownsp WHERE (((msalesbrkdownsp.SPTotal)>0 AND ((msalesbrkdownsp.SDATE)='" & Format(Date, "yyyy-mm-dd") & "'));"
    Else
        Qry = "SELECT msalesbrkdownsvr.* from msalesbrkdownsvr WHERE (((msalesbrkdownsvr.SPTotal)>0 AND ((msalesbrkdownsvr.SDATE)='" & Format(Date, "yyyy-mm-dd") & "'));"
    End If
    With rsTSales
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            If chkSel.Value = 0 Then
                .MoveFirst
                flxSPDay.TextMatrix(2, 1) = Format(!WALKIN, "#,##0.00")
                flxSPDay.TextMatrix(2, 2) = Format(!WHLSALE, "#,##0.00")
                flxSPDay.TextMatrix(2, 3) = Format(!RBLTPLAN, "#,##0.00")
                flxSPDay.TextMatrix(2, 4) = Format(!LGCMKTG, "#,##0.00")
                flxSPDay.TextMatrix(2, 5) = Format(!LGCFIN, "#,##0.00")
                flxSPDay.TextMatrix(2, 6) = Format(!WHLACCT, "#,##0.00")
                flxSPDay.TextMatrix(2, 7) = Format(!EMPACCT, "#,##0.00")
                flxSPDay.TextMatrix(2, 8) = Format(!CORTACCT, "#,##0.00")
                flxSPDay.TextMatrix(2, 9) = Format(!SRETURN, "#,##0.00")
                flxSPDay.TextMatrix(2, 10) = Format(!AdjDebit, "#,##0.00")
                flxSPDay.TextMatrix(2, 11) = Format(!AdjCredit, "#,##0.00")
                flxSPDay.TextMatrix(2, 12) = Format(!SOTHERS, "#,##0.00")
                flxSPDay.TextMatrix(2, 13) = Format(!SPTotal, "#,##0.00")
            Else
                .MoveFirst
                flxSVRDay.TextMatrix(2, 1) = Format(!WALKIN, "#,##0.00")
                flxSVRDay.TextMatrix(2, 2) = Format(!LGCMKTG, "#,##0.00")
                flxSVRDay.TextMatrix(2, 3) = Format(!LGCFIN, "#,##0.00")
                flxSVRDay.TextMatrix(2, 4) = Format(!EMPACCT, "#,##0.00")
                flxSVRDay.TextMatrix(2, 5) = Format(!CORTACCT, "#,##0.00")
                flxSVRDay.TextMatrix(2, 6) = Format(!AdjDebit, "#,##0.00")
                flxSVRDay.TextMatrix(2, 7) = Format(!AdjCredit, "#,##0.00")
                flxSVRDay.TextMatrix(2, 8) = Format(SOTHERS, "#,##0.00")
                flxSVRDay.TextMatrix(2, 9) = Format(!SVRTOTAL, "#,##0.00")
            End If
        End If
        .Close
    End With
End Sub

Private Sub Timer3_Timer()
    LoadDaySales
    Timer3.Enabled = False
End Sub
