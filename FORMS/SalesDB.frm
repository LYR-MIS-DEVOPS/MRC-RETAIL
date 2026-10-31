VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SalesDB 
   Caption         =   "Merchandise Inventory"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   14310
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   9120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid FlxRpt 
      Height          =   7335
      Left            =   120
      TabIndex        =   29
      Tag             =   "xyf"
      Top             =   1080
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12938
      _Version        =   393216
      Cols            =   20
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      AllowUserResizing=   1
      FormatString    =   $"SalesDB.frx":0000
   End
   Begin TabDlg.SSTab tabRepair 
      Height          =   2415
      Left            =   4440
      TabIndex        =   17
      Top             =   9120
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SalesDB.frx":01DC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRefNo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblRefDate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblRecFrom"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblRecBy"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtRefNo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtRecFrom"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtRecBy"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   2880
         TabIndex        =   26
         Top             =   1680
         Width           =   3615
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   165
            Width           =   1440
            _ExtentX        =   2540
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
            Image           =   "SalesDB.frx":01F8
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   375
            Left            =   1920
            TabIndex        =   28
            Top             =   165
            Width           =   1440
            _ExtentX        =   2540
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
            Image           =   "SalesDB.frx":0792
            cBack           =   -2147483633
         End
      End
      Begin VB.TextBox txtRecBy 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   25
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txtRecFrom 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   23
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txtDate 
         Height          =   300
         Left            =   5040
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtRefNo 
         Height          =   300
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblRecBy 
         Alignment       =   1  'Right Justify
         Caption         =   "REC. BY:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label lblRecFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "REC. FROM:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   975
         Width           =   1215
      End
      Begin VB.Label lblRefDate 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE:"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   383
         Width           =   1215
      End
      Begin VB.Label lblRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. NO.:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   383
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   135
      Left            =   120
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   8520
      Visible         =   0   'False
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   9000
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   10680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   580
      Left            =   120
      TabIndex        =   8
      Tag             =   "XY"
      Top             =   480
      Width           =   14055
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   101187585
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         ToolTipText     =   "Execute"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "SalesDB.frx":0D2C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         ToolTipText     =   "Refresh"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "SalesDB.frx":12C6
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   12
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   101187585
         CurrentDate     =   39322
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL RECORDS: 0"
         Height          =   255
         Left            =   11760
         TabIndex        =   30
         Tag             =   "XYF"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxCV 
      Height          =   1215
      Left            =   360
      TabIndex        =   7
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "CV No   |<Amount    |<Date        |<Remarks"
   End
   Begin CRVIEWERLibCtl.CRViewer vwrRpt 
      Height          =   7935
      Left            =   120
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   480
      Width           =   14055
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
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   9240
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   8760
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   9240
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   8760
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
      Image           =   "SalesDB.frx":1860
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
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
      Image           =   "SalesDB.frx":1DFA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
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
      Image           =   "SalesDB.frx":2394
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Save to Excel"
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
      Image           =   "SalesDB.frx":292E
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
      Image           =   "SalesDB.frx":2EC8
      cBack           =   -2147483633
   End
   Begin VB.Label lblRemarks 
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   10320
      Visible         =   0   'False
      Width           =   3735
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
      Left            =   120
      TabIndex        =   6
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14550
      m_MDI_Height    =   9330
   End
   Begin VB.Menu mnuProd 
      Caption         =   "Product Entry"
      Visible         =   0   'False
      Begin VB.Menu mnuProdEntry 
         Caption         =   "Show Product Entry Form"
      End
      Begin VB.Menu mnuPrcQuote 
         Caption         =   "Show Price Quotation"
      End
   End
   Begin VB.Menu mnuForRepair 
      Caption         =   "For Repair"
      Visible         =   0   'False
      Begin VB.Menu mnuHistory 
         Caption         =   "Show History"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepaired 
         Caption         =   "Update as ""REPAIRED UNIT"""
      End
      Begin VB.Menu mnuCM 
         Caption         =   "Issued with Credit Memo (CM)"
      End
   End
End
Attribute VB_Name = "SalesDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer, ActRow As Integer, Qry As String
Dim fRow As Integer, fCol As Integer, TCol As Integer
Public mySrc As String

Public Sub ShowFilteredRecords()
    If mySrc = "Merchandise Inventory" Then
        Qry = "SELECT productline.CLASSIFICATION, productline.PRODCODE, productline.BRAND, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.SERIALNO, items.COLOR, items.QTYONHAND, items.DADR, items.DATEREC, DATEDIFF(NOW(),DATEREC) AS AGE, items.SOURCE, items.`STATUS`, productline.PRODINIT, items.GRPCODE, productline.PRODID " & _
              "From items INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
              txtCriteria.Text & " AND items.LOCATION <> 'NOT YET PRINTED' AND (items.GRPCODE <= 7 AND items.GRPCODE <> 5) Order By GRPCODE, items.`STATUS`, PRODINIT, productline.CLASSIFICATION, productline.BRAND, productline.MODEL, items.COLOR, items.DATEREC;"
    
    ElseIf mySrc = "Price List - Brand New" Then
        'Qry = "SELECT ITEMS.DATEREC, ITEMS.DADR, PRODUCTLINE.MODEL, [BRAND]+' '+[DESCRIPTION] AS [DESC], ITEMS.SERIALNO, ITEMS.QTYONHAND, DateDiff('d',[DATEREC],Now()) AS AGE, ITEMS.STATUS, PRODUCTLINE.PRICELISTID, PRODUCTLINE.PRODID " & _
              "FROM ITEMS INNER JOIN PRODUCTLINE ON ITEMS.PRODID = PRODUCTLINE.PRODID WHERE (((ITEMS.GRPCODE)=1) AND ((ITEMS.LOCATION)<>'NOT YET PRINTED')) ORDER BY PRODUCTLINE.PRODTYPE, PRODUCTLINE.BRAND;"
        Qry = "SELECT DATE_FORMAT(DATEREC,'%c/%d/%Y'), items.DADR, productline.MODEL, CONCAT(BRAND,' ',DESCRIPTION) AS MYDESC, items.SERIALNO, items.QTYONHAND, DATEDIFF(NOW(),DATEREC) AS AGE, ITEMS.STATUS, PRODUCTLINE.PRICELISTID, PRODUCTLINE.PRODID " & _
              "From items INNER JOIN productline ON (items.PRODID = productline.PRODID) WHERE (items.LOCATION <>'NOT YET PRINTED' AND ITEMS.GRPCODE<=4) ORDER BY productline.PRODTYPE ASC, productline.BRAND ASC;"
    ElseIf mySrc = "Service Unit" Then
        Qry = "SELECT productline.CLASSIFICATION, productline.PRODCODE, productline.BRAND, productline.MODEL, IF(productline.PRODINIT = 'MC' OR productline.PRODINIT = 'FR', '', productline.DESCRIPTION) AS MYDESC, items.SERIALNO, items.COLOR, items.QTYONHAND, items.DADR, items.DATEREC, DATEDIFF(NOW(),DATEREC) AS AGE, items.REMARKS, items.`STATUS`, productline.PRODINIT, items.GRPCODE, productline.PRODID " & _
              "From items INNER JOIN productline ON (items.PRODID = productline.PRODID) " & _
              txtCriteria.Text & " AND items.LOCATION <> 'NOT YET PRINTED' AND items.GRPCODE = 5 Order By GRPCODE, items.`STATUS`, PRODINIT, productline.CLASSIFICATION, productline.BRAND, productline.MODEL, items.COLOR, items.DATEREC;"
    End If
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsRPT.RecordCount = 0 Then
        FlxRpt.Clear
        FlxRpt.Rows = 2
        FlxRpt.Cols = TCol
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set FlxRpt.DataSource = rsRPT
    ActRow = 1
    Call RXResizer1_ResizeComplete

End Sub


Private Sub cmdCancel_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Cancel transaction?", vbQuestion + vbYesNo, "Update as REPAIRED UNIT")
    If Ans = vbNo Then Exit Sub
    tabRepair.Visible = False
    tabRepair.Top = 15000
    For i = 2 To 4
        cmdTrans(i).Enabled = True
    Next i
    Frame1.Enabled = True
    FlxRpt.Enabled = True
End Sub


Private Sub cmdRefresh_Click()
    Dim Qry As String, i As Long
    Dim rsInv As New ADODB.Recordset
    Dim rsC As New ADODB.Recordset
    
    
    cmdTrans(2).Enabled = False
    cmdTrans(3).Enabled = False
    cmdTrans(4).Enabled = False
    Frame1.Enabled = False
    Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.REMARKS AS Classif, spitr.TRANSDATE, spitr.REFERENCE, spitr.OUT, spareparts.COST, spitr.PRICE, spareparts.SUPPLIERCODE, spitr.REMARKS, spitr.Source " & _
          "From spitr INNER JOIN spareparts ON (spitr.SPAREPARTID = spareparts.SPAREPARTID) " & _
          "Where spitr.TRANSTYPE = 'OUT' AND spitr.TRANSDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND spitr.TRANSDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' AND  spitr.REFERENCE LIKE 'SI%' Order By spareparts.PARTNUMBER, spitr.TRANSDATE;"
              
    Pbar.Value = 0
    Pbar.Visible = True
    rsInv.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsInv.RecordCount > 0 Then
        FlxRpt.Rows = 1
        Pbar.Max = rsInv.RecordCount * 2
        Label2.Caption = "TOTAL RECORDS: " & rsInv.RecordCount
        rsInv.MoveFirst
        Do While Not rsInv.EOF
            FlxRpt.AddItem vbTab & rsInv!PartNumber & vbTab & rsInv!Description & vbTab & rsInv!Classif & vbTab & Format(rsInv!TRANSDATE, "mm/dd/yyyy") & vbTab & rsInv!REFERENCE & vbTab & rsInv!OUT & vbTab & _
            Format(rsInv!Cost, "#,##0.00") & vbTab & Format(rsInv!Price, "#,##0.00") & vbTab & rsInv!SUPPLIERCODE & vbTab & rsInv!Remarks & vbTab & rsInv!Source
            rsInv.MoveNext
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Loop
    Else
        FlxRpt.Rows = 1
        FlxRpt.Rows = 2
        MsgBox "No records found!", vbExclamation + vbOKOnly, "Sales DB"
        rsInv.Close
        cmdTrans(2).Enabled = True
        cmdTrans(3).Enabled = True
        cmdTrans(4).Enabled = True
        Frame1.Enabled = True
        
        Exit Sub
    End If
    rsInv.Close
    
    'Show Customer
    '1:PART NUMBER, 2:DESCRIPTION, 3:CLASSIFICATION 4:DATE SOLD, 5:SALES INV#., 6:QTY., 7:COST W/O VAT, 8:UNIT PRICE, 9:SUPPLIER, 10:TRANS. TYPE,
    '11:CUSTOMER/DEALER, 12:ADDRESS, 13:CITY/TOWN, 14:PROVINCE, 15:REGION, 16:BRANCHES, 17:CUSTOMER TYPE, 18:FRONTLINE 19:CUST.ID
    With FlxRpt
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, 10) = "CREDIT" Then
                Qry = "SELECT customer.CUSTOMERNAME, customer.ADDRESS, customer.CITYTOWN, customer.PROVINCE, customer.REGION, customer.BRANCHES, customer.CTYPE, spinvoice.SPINVNO, customer.CUSTOMERID " & _
                      "From spinvoice INNER JOIN customer ON (spinvoice.CUSTOMERID = customer.CUSTOMERID) " & _
                      "Where spinvoice.TRANSNO = '" & .TextMatrix(i, 5) & "';"
                rsC.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If rsC.RecordCount > 0 Then
                    .TextMatrix(i, 11) = rsC!CustomerName
                    .TextMatrix(i, 12) = IIf(IsNull(rsC!Address), "", rsC!Address)
                    .TextMatrix(i, 13) = IIf(IsNull(rsC!CITYTOWN), "", rsC!CITYTOWN)
                    .TextMatrix(i, 14) = IIf(IsNull(rsC!PROVINCE), "", rsC!PROVINCE)
                    .TextMatrix(i, 15) = IIf(IsNull(rsC!REGION), "", rsC!REGION)
                    .TextMatrix(i, 16) = IIf(IsNull(rsC!BRANCHES), "", rsC!BRANCHES)
                    .TextMatrix(i, 17) = IIf(IsNull(rsC!Ctype), "", rsC!Ctype)
                    .TextMatrix(i, 19) = rsC!CUSTOMERID
                End If
                rsC.Close
            ElseIf .TextMatrix(i, 10) = "CASH" Then
                .TextMatrix(i, 12) = IIf(.TextMatrix(i, 12) <> "", UCase$(.TextMatrix(i, 12)), "TAGUM CITY")
                .TextMatrix(i, 17) = GetFieldValue("SELECT spinvoice.CTYPE From spinvoice Where spinvoice.TRANSNO = '" & .TextMatrix(i, 5) & "';")
                .TextMatrix(i, 17) = IIf(.TextMatrix(i, 17) <> "", UCase$(.TextMatrix(i, 17)), "WALK-IN")
                
                If Val(.TextMatrix(i, 19)) > 0 Then
                    .TextMatrix(i, 11) = GetFieldValue("SELECT spinvoice.CHARGETO From spinvoice Where spinvoice.TRANSNO = '" & .TextMatrix(i, 5) & "';")
                    Qry = "SELECT customer.CUSTOMERNAME, customer.ADDRESS, customer.CITYTOWN, customer.PROVINCE, customer.REGION, customer.BRANCHES, customer.CTYPE, customer.CUSTOMERID " & _
                          "From customer Where customer.CUSTOMERID = " & Val(.TextMatrix(i, 19)) & " AND customer.CUSTOMERNAME = '" & .TextMatrix(i, 11) & "';"
                    rsC.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If rsC.RecordCount > 0 Then
                        If .TextMatrix(i, 11) <> "" And .TextMatrix(i, 11) = rsC!CustomerName Then
                            .TextMatrix(i, 11) = rsC!CustomerName
                            .TextMatrix(i, 12) = IIf(IsNull(rsC!Address), "", rsC!Address)
                            .TextMatrix(i, 13) = IIf(IsNull(rsC!CITYTOWN), "", rsC!CITYTOWN)
                            .TextMatrix(i, 14) = IIf(IsNull(rsC!PROVINCE), "", rsC!PROVINCE)
                            .TextMatrix(i, 15) = IIf(IsNull(rsC!REGION), "", rsC!REGION)
                            .TextMatrix(i, 16) = IIf(IsNull(rsC!BRANCHES), "", rsC!BRANCHES)
                            .TextMatrix(i, 17) = IIf(IsNull(rsC!Ctype), "", rsC!Ctype)
                            .TextMatrix(i, 19) = rsC!CUSTOMERID
                        End If
                    End If
                    rsC.Close
                End If
            End If
            .TextMatrix(i, 9) = GetFieldValue("SELECT supplier.SUPPLIERNAME From supplier Where supplier.SUPPLIERCODE = '" & .TextMatrix(i, 9) & "';")
            .TextMatrix(i, 18) = GetFieldValue("SELECT spinvoice.FRONTLINER From spinvoice Where spinvoice.TRANSNO = '" & .TextMatrix(i, 5) & "';")
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
    End With
    Pbar.Visible = False
    cmdTrans(2).Enabled = True
    cmdTrans(3).Enabled = True
    cmdTrans(4).Enabled = True
    Frame1.Enabled = True
    MsgBox "Sales Database Report successfully loaded!", vbInformation + vbOKOnly, "Sales DB"
    

End Sub


Private Sub cmdSave_Click()
    Dim MyGrp As Integer, Ans As Integer, i As Integer
    Dim vMsg As String
    
    'Validate transaction
    vMsg = ""
    If txtRefNo.Text = "" Then vMsg = "Incomplete entry on field " & lblRefNo.Caption
    If Not IsDate(txtDate.Text) Then vMsg = vMsg & vbCrLf & "Invalid entry on field DATE"
    If txtRecBy.Text = "" Then vMsg = vMsg & vbCrLf & "Please identify source of documents"
    
    If vMsg <> "" Then
        MsgBox vMsg, vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Save transaction?", vbQuestion + vbYesNo, cmdSave.Tag)
    If Ans = vbNo Then Exit Sub
    
    MyGrp = GetGroupCode(FlxRpt.TextMatrix(ActRow, 6))
    If cmdSave.Tag = "REPAIRED" Then
        LYRcon.Execute "INSERT INTO invtreg(TRANSID, REFERENCE, TRANSDATE, TRANSTYPE, ITEMID, `REC`, ORDERID, SOURCE, CBOWNER, SRCTYPE, UPDATEDBY, LASTUPDATE, REMARKS) " & _
              "VALUES(0, 'REF#:" & txtRefNo.Text & "', '" & Format(DateValue(txtDate.Text), "yyyy-mm-dd") & "', 'SERVICE', " & Val(FlxRpt.TextMatrix(ActRow, 9)) & ", 1, " & Val(FlxRpt.TextMatrix(ActRow, 10)) & ", '" & txtRecFrom.Text & "', '" & txtRecBy.Text & "', ' REPAIRED UNIT" & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & FlxRpt.TextMatrix(ActRow, 6) & "');"
        LYRcon.Execute "UPDATE items SET `STATUS` = '" & FlxRpt.TextMatrix(ActRow, 6) & "', GRPCODE = " & MyGrp & " Where items.ItemID = " & Val(FlxRpt.TextMatrix(ActRow, 9)) & ";"
        MsgBox "Repaired transaction successfully saved!", vbInformation + vbOKOnly, "Save"
    ElseIf cmdSave.Tag = "For CM" Then
        LYRcon.Execute "INSERT INTO invtreg(TRANSID, REFERENCE, TRANSDATE, TRANSTYPE, ITEMID, `OUT`, ORDERID, SOURCE, CBOWNER, SRCTYPE, UPDATEDBY, LASTUPDATE, REMARKS) " & _
              "VALUES(0, 'CM#:" & txtRefNo.Text & "', '" & Format(DateValue(txtDate.Text), "yyyy-mm-dd") & "', 'RETURN', " & Val(FlxRpt.TextMatrix(ActRow, 9)) & ", 1, " & Val(FlxRpt.TextMatrix(ActRow, 10)) & ", '" & txtRecFrom.Text & "', '" & txtRecBy.Text & "', ' RETURN UNIT W/ CM" & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & FlxRpt.TextMatrix(ActRow, 6) & "');"
        LYRcon.Execute "UPDATE items SET `STATUS` = 'RETURNED', GRPCODE = 12 Where items.ItemID = " & Val(FlxRpt.TextMatrix(ActRow, 9)) & ";"
        MsgBox "Returned unit with CM successfully saved!", vbInformation + vbOKOnly, "Save"
    End If
    If ActRow = 1 Then
        FlxRpt.Rows = 1
        FlxRpt.Rows = 2
    Else
        FlxRpt.RemoveItem ActRow
    End If
    tabRepair.Visible = False
    tabRepair.Top = 15000
    For i = 2 To 4
        cmdTrans(i).Enabled = True
    Next i
    Frame1.Enabled = True
    FlxRpt.Enabled = True

End Sub

Function GetGroupCode(MyStat As String) As Integer
    
    Select Case MyStat
        Case "BRAND NEW"
            GetGroupCode = 1
        Case "CLASS B"
            GetGroupCode = 2
        Case "REPO."
            GetGroupCode = 3
        Case "ACCESSORIES"
            GetGroupCode = 4
        Case "SERVICE UNIT"
            GetGroupCode = 5
        Case "PROMO"
            GetGroupCode = 6
        Case "DEPOSITED"
            GetGroupCode = 7
        Case "PENDING (SOLD)"
            GetGroupCode = 10
        Case "FOR REPAIR"
            GetGroupCode = 13
        Case Else
            GetGroupCode = 0
    End Select
End Function

Private Sub cmdTrans_Click(Index As Integer)
    Dim rsPrn As New ADODB.Recordset
    'Dim crMerch As New RptInvMerchnew
    'Dim crSvc As New RptInvSvcUnit
    Dim MyQry As String, i As Long, j As Long
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
        Case 1  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            'Disbursement.Show
        Case 2  'Filter
            MsgBox "Filtering not yet available!", vbInformation + vbOKOnly, "SALES DB"
            Exit Sub
            If mySrc = "Merchandise Inventory" Then
                For i = 0 To 2
                     With frmFilter.cboLookIn(i)
                       .AddItem "PRODUCT TYPE"
                       .AddItem "CLASSIFICATION"
                       .AddItem "STATUS"
                       .AddItem "BRAND"
                       .AddItem "DESCRIPTION"
                       .AddItem "MODEL"
                       .AddItem "SERIAL NO."
                       .AddItem "REF. NO."
                       .AddItem "DATE RECEIVED"
                       .AddItem "AGE"
                       .AddItem "SOURCE"
                     End With
                Next i
                frmFilter.fldNames = "productline.PRODINIT*productline.CLASSIFICATION*items.STATUS*productline.BRAND*productline.DESCRIPTION*productline.MODEL*items.SERIALNO*items.DADR*items.DATEREC*DATEDIFF(NOW(),DATEREC)*items.SOURCE"
                'FRMFILTER.fldNames = "Disbursement.CVNUMBER*Disbursement.DISBURSETYPE*Disbursement.ACCTDESC*Disbursement.PAYEE*Disbursement.PARTICULARS*Disbursement.DEPARTMENT"
                frmFilter.fldNdx = "9" 'start w/ 0
                frmFilter.fldDates = "8"
                frmFilter.fldSupplier = 11  'Used no > index to exclude field supplier
                frmFilter.MySource = "Merchandise Inventory"
                frmFilter.Show vbModal
            ElseIf mySrc = "Service Unit" Then
                For i = 0 To 2
                     With frmFilter.cboLookIn(i)
                       .AddItem "PRODUCT TYPE"
                       .AddItem "CLASSIFICATION"
                       .AddItem "STATUS"
                       .AddItem "BRAND"
                       .AddItem "DESCRIPTION"
                       .AddItem "MODEL"
                       .AddItem "SERIAL NO."
                       .AddItem "REF. NO."
                       .AddItem "DATE RECEIVED"
                       .AddItem "AGE"
                       .AddItem "ASSIGNED TO"
                     End With
                Next i
                frmFilter.fldNames = "productline.PRODINIT*productline.CLASSIFICATION*items.STATUS*productline.BRAND*productline.DESCRIPTION*productline.MODEL*items.SERIALNO*items.DADR*items.DATEREC*DATEDIFF(NOW(),DATEREC)*items.REMARKS"
                'FRMFILTER.fldNames = "Disbursement.CVNUMBER*Disbursement.DISBURSETYPE*Disbursement.ACCTDESC*Disbursement.PAYEE*Disbursement.PARTICULARS*Disbursement.DEPARTMENT"
                frmFilter.fldNdx = "9" 'start w/ 0
                frmFilter.fldDates = "8"
                frmFilter.fldSupplier = 11  'Used no > index to exclude field supplier
                frmFilter.MySource = "Service Unit"
                frmFilter.Show vbModal
            End If
        Case 3  'SAVE TO EXCEL
            Dim XLapp As Excel.Application
            Dim wbkMVR As Excel.Workbook
            Dim wksLineA1 As Excel.Worksheet
 
            Dim sSaveAsFileName
            
            With dlgDEMO
                .CancelError = True
                .InitDir = mstrLastDir
                .Flags = cdlOFNHideReadOnly
                .filename = ""
                .Filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
                .ShowOpen
                sSaveAsFileName = .filename
            End With
            
 
            Set XLapp = CreateObject("Excel.Application")
            XLapp.Visible = True
            XLapp.Workbooks.Add
            Set wbkMVR = XLapp.ActiveWorkbook
            Set wksLineA1 = wbkMVR.Sheets(1)

    'Show Customer
    '1:PART NUMBER, 2:DESCRIPTION, 3:CLASSIFICATION 4:DATE SOLD, 5:SALES INV#., 6:QTY., 7:COST W/O VAT, 8:UNIT PRICE, 9:SUPPLIER, 10:TRANS. TYPE,
    '11:CUSTOMER/DEALER, 12:ADDRESS, 13:CITY/TOWN, 14:PROVINCE, 15:REGION, 16:BRANCHES, 17:CUSTOMER TYPE, 18:FRONTLINE 18:CUST.ID

            wksLineA1.Name = "Sales"
                wksLineA1.Cells(1, 1).Value = "PARTNUMBER"
                wksLineA1.Cells(1, 2).Value = "DESCRIPTION"
                wksLineA1.Cells(1, 3).Value = "CLASSIFICATION"
                wksLineA1.Cells(1, 4).Value = "DATE"
                wksLineA1.Cells(1, 5).Value = "INVOICE"
                wksLineA1.Cells(1, 6).Value = "QTY"
                wksLineA1.Cells(1, 7).Value = "COST"
                wksLineA1.Cells(1, 8).Value = "PRICE"
                wksLineA1.Cells(1, 9).Value = "SUPPLIER"
                wksLineA1.Cells(1, 10).Value = "TRANSTYPE"
                wksLineA1.Cells(1, 11).Value = "CUSTOMER"
                wksLineA1.Cells(1, 12).Value = "ADDRESS"
                wksLineA1.Cells(1, 13).Value = "CITY/TOWN"
                wksLineA1.Cells(1, 14).Value = "PROVINCE"
                wksLineA1.Cells(1, 15).Value = "REGION"
                wksLineA1.Cells(1, 16).Value = "BRANCHES"
                wksLineA1.Cells(1, 17).Value = "CUSTTYPE"
                wksLineA1.Cells(1, 18).Value = "FRONTLINER"
                With FlxRpt
                    Pbar.Value = 0
                    Pbar.Max = .Rows - 1
                    For i = 1 To .Rows - 1
                        For j = 1 To 18
                            wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                        Next j
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                    Next i
                End With
                XLapp.DisplayAlerts = False
                MyQry = GetFieldValue("SELECT pcuser.Password From pcuser Where pcuser.EmpId = 4;")
                wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:=MyQry, WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
 
                wbkMVR.Close SaveChanges:=False
                XLapp.Quit
 
                Set XLapp = Nothing
                Set wbkMVR = Nothing
                Set wksLineA1 = Nothing
                MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, "Sales DB"
        Case 4  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
            'Disbursement.Show
    End Select
End Sub

Private Sub dFrom_CloseUp()
    'If cmdNext.Tag = "1" Then
    '    Call RefreshRecord(1)
    'End If
End Sub

Private Sub dTo_CloseUp()
    'If dTo.Value < dFrom.Value Then
    '    MsgBox "Please choose a date not less than " & dFrom.Value, vbExclamation + vbOKOnly, "Validation Error"
    '    dTo.Value = dFrom.Value
    'End If
End Sub

Private Sub flxRpt_EnterCell()
    ActRow = FlxRpt.Row
End Sub

Private Sub flxRpt_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If mySrc = "Merchandise Inventory" Then
            Call PopupMenu(mnuProd)
        ElseIf mySrc = "For Repair" Then
            Call PopupMenu(mnuForRepair)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dTo.Value = Date
    cmdNext.Tag = "1"
    MotoRep.sBar.Panels(6).Text = "SALES DATABASE REPORT"
    Me.Caption = mySrc
    BtnIndx = 1
    tmrFloat.Enabled = True
    Call PrintPreView(False)
End Sub

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
        Frame1.Visible = False
        FlxRpt.Visible = False
        vwrRpt.Visible = True
    Else
        vwrRpt.Visible = False
        Frame1.Visible = True
        FlxRpt.Visible = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub mnuCM_Click()
    Dim MyAcc As String, Qry As String
    Dim i As Integer
    
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=137;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    For i = 2 To 4
        cmdTrans(i).Enabled = False
    Next i
    Frame1.Enabled = False
    FlxRpt.Enabled = False
    lblRefNo.Caption = "CM NO."
    lblRecBy.Caption = "REMARKS:"
    
    With tabRepair
        .Caption = ""
        .Top = Me.Height / 2 - .Height
        .Left = (Me.Width - .Width) / 2
        .Visible = True
    End With
    cmdSave.Tag = "For CM"

End Sub

Private Sub mnuHistory_Click()
    InvHistory.ItemID = Val(FlxRpt.TextMatrix(ActRow, 9))
    InvHistory.Show vbModal
End Sub

Private Sub mnuPrcQuote_Click()
    PrcQuotation.mySrc = "INV. REPORT"
    PrcQuotation.FtxtModel.Text = FlxRpt.TextMatrix(FlxRpt.Row, 4)
    PrcQuotation.ftxtProdID.Text = FlxRpt.TextMatrix(FlxRpt.Row, 16)
    Call PrcQuotation.ShowRecord
    PrcQuotation.Show vbModal
End Sub

Private Sub mnuProdEntry_Click()
    Call InvProduct.ShowRecord(FlxRpt.TextMatrix(FlxRpt.Row, 16))
    InvProduct.cmdTrans(3).Tag = "MERCHANDISE"
    InvProduct.Show vbModal
End Sub

Private Sub mnuRepaired_Click()
    Dim MyAcc As String, Qry As String
    Dim i As Integer
    
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=137;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If

    For i = 2 To 4
        cmdTrans(i).Enabled = False
    Next i
    Frame1.Enabled = False
    FlxRpt.Enabled = False
    lblRefNo.Caption = "REF NO."
    lblRecBy.Caption = "REC. BY:"
    With tabRepair
        .Caption = ""
        .Top = Me.Height / 2 - .Height
        .Left = (Me.Width - .Width) / 2
        .Visible = True
    End With
    cmdSave.Tag = "REPAIRED"

End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, Qry As String
    
    With FlxRpt
        If mySrc = "Merchandise Inventory" Then
            '                   CLASSIFICATION, PROD. CODE,  BRAND,    MODEL,   MYDESC,        SERIALNO,    COLOR,QTYONHAND, DADR,     DATEREC, DATEDIFF(NOW(),DATEREC) AS AGE, items.SOURCE, items.`STATUS`, productline.PRODINIT, items.GRPCODE "
            .FormatString = "  |<CLASSIFICATION |<PROD. CODE|<BRAND |<MODEL  |<DESCRIPTION |< SERIAL NO. |<COLOR  |^QTY. |<REF. NO. |^DATE REC.  |^AGE|<SOURCE|<STATUS|^TYPE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.15
            .ColWidth(2) = (.Width - 400) * 0.1
            .ColWidth(3) = (.Width - 400) * 0.1
            .ColWidth(4) = (.Width - 400) * 0.15
            .ColWidth(5) = (.Width - 400) * 0.15
            .ColWidth(6) = (.Width - 400) * 0.08
            .ColWidth(7) = (.Width - 400) * 0.08
            .ColWidth(8) = (.Width - 400) * 0.05
            .ColWidth(9) = (.Width - 400) * 0.05
            .ColWidth(10) = (.Width - 400) * 0.05
            .ColWidth(11) = (.Width - 400) * 0.05
            .ColWidth(12) = (.Width - 400) * 0.15
            .ColWidth(13) = (.Width - 400) * 0.1
            .ColWidth(14) = (.Width - 400) * 0.03
            .ColWidth(15) = 0
            .ColWidth(16) = 0
            'For i = 1 To .Rows - 1
            '    .TextMatrix(i, 6) = Format(.TextMatrix(i, 6), "mm/dd/yyyy")
            'Next i
        '        Qry = "DATEREC,'%c/%d/%Y'), items.DADR, productline.MODEL, CONCAT(BRAND,' ',CLASSIFICATION) AS MYDESC, items.SERIALNO, items.COLOR, items.QTYONHAND, DATEDIFF(NOW(),DATEREC) AS AGE, ITEMS.STATUS, PRODUCTLINE.PRICELISTID, PRODUCTLINE.PRODID "
        ElseIf mySrc = "Price List - Brand New" Then
            .FormatString = "   |^DATE EFF. |^PRICE REF#. |<MODEL |<BRAND - CLASSIFICATION |< SERIAL NO.  |<COLOR |^QTY. |^AGE|<STATUS "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.1
            .ColWidth(2) = (.Width - 400) * 0.1
            .ColWidth(3) = (.Width - 400) * 0.15
            .ColWidth(4) = (.Width - 400) * 0.25
            .ColWidth(5) = (.Width - 400) * 0.1
            .ColWidth(6) = (.Width - 400) * 0.1
            .ColWidth(7) = (.Width - 400) * 0.08
            .ColWidth(8) = (.Width - 400) * 0.08
            .ColWidth(9) = (.Width - 400) * 0.1
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            Pbar.Top = .Top + .Height
            Pbar.Left = .Left
            Pbar.Width = .Width
        ElseIf mySrc = "Service Unit" Then
            .FormatString = "  |<CLASSIFICATION |<PROD. CODE|<BRAND |<MODEL  |<DESCRIPTION |< SERIAL NO. |<COLOR  |^QTY. |<REF. NO. |^DATE REC.  |^AGE|<ASSIGNED TO|<STATUS|^TYPE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.15
            .ColWidth(2) = (.Width - 400) * 0.1
            .ColWidth(3) = (.Width - 400) * 0.1
            .ColWidth(4) = (.Width - 400) * 0.15
            .ColWidth(5) = (.Width - 400) * 0.15
            .ColWidth(6) = (.Width - 400) * 0.08
            .ColWidth(7) = (.Width - 400) * 0.08
            .ColWidth(8) = (.Width - 400) * 0.05
            .ColWidth(9) = (.Width - 400) * 0.05
            .ColWidth(10) = (.Width - 400) * 0.05
            .ColWidth(11) = (.Width - 400) * 0.05
            .ColWidth(12) = (.Width - 400) * 0.15
            .ColWidth(13) = (.Width - 400) * 0.1
            .ColWidth(14) = (.Width - 400) * 0.03
            .ColWidth(15) = 0
            .ColWidth(16) = 0
            
        ElseIf mySrc = "For Repair" Then
            '1-SVC DATE, 2-SVC REQ#, 3-BRAND, 4-MODEL, 5-SERIAL, 6-STATUS, 7-LOCATION, 8-REMARKS, 9-ITEMID, 10-SVCREQID
            .FormatString = "  |<SVC DATE |^SVC.REQ# |<BRAND  |<MODEL |<SERIAL NO. |<STATUS |<LOCATION  |<REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.08
            .ColWidth(2) = (.Width - 400) * 0.08
            .ColWidth(3) = (.Width - 400) * 0.12
            .ColWidth(4) = (.Width - 400) * 0.15
            .ColWidth(5) = (.Width - 400) * 0.1
            .ColWidth(6) = (.Width - 400) * 0.08
            .ColWidth(7) = (.Width - 400) * 0.2
            .ColWidth(8) = (.Width - 400) * 0.2
            .ColWidth(9) = 0
            .ColWidth(10) = 0
        End If
    End With
    Label2.Height = cmdRefresh.Height
    Label2.Top = cmdRefresh.Top
    'txtFind.Top = cmdFind.Top
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer, Qry As String
    
    For i = 0 To 4
        cmdTrans(i).Enabled = False
    Next i
    With FlxRpt
        If mySrc = "Price List - Brand New" Then
            Pbar.Value = 0
            Pbar.Max = .Rows - 1
            Pbar.Visible = True
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 10)) > 0 Then
                    Qry = "SELECT PRICELIST.EFFECTDATE From PRICELIST WHERE (PRICELISTID=" & .TextMatrix(i, 10) & ");"
                    .TextMatrix(i, 1) = GetFieldValue(Qry)
                    Qry = "SELECT PRICELIST.REFNO From PRICELIST WHERE (PRICELISTID=" & .TextMatrix(i, 10) & ");"
                    .TextMatrix(i, 2) = GetFieldValue(Qry)
                Else
                    .TextMatrix(i, 1) = ""
                    .TextMatrix(i, 2) = ""
                End If
                Pbar.Value = Pbar.Value + 1
                DoEvents
            Next i
            Pbar.Visible = False
        End If
    End With
    For i = 0 To 4
        cmdTrans(i).Enabled = True
    Next i
    Timer1.Enabled = False
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 5 Then     'last index + 1
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
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 2 'Index of Button view grid
            For i = 2 To 4
                cmdTrans(i).Top = -375
                cmdTrans(i).Left = IIf(i = 2, 0, cmdTrans(2).Width * (i - 2))
                cmdTrans(i).Visible = True
            Next i
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

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890-/"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtRecFrom.SetFocus
End Sub


Private Sub txtRecBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub


Private Sub txtRecFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then txtRecBy.SetFocus

End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then txtDate.SetFocus
End Sub
