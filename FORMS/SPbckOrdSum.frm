VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SPbckOrdSum 
   Caption         =   "Back Order Summary"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabSort 
      Height          =   2655
      Left            =   7440
      TabIndex        =   27
      Top             =   8520
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "You can Sort records upto 4 fields"
      TabPicture(0)   =   "SPbckOrdSum.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSortCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSortOK"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   3735
         Begin VB.CheckBox chkSort 
            Caption         =   "Ascending"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1185
            Width           =   1095
         End
         Begin VB.CheckBox chkSort 
            Caption         =   "Ascending"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   870
            Width           =   1095
         End
         Begin VB.CheckBox chkSort 
            Caption         =   "Ascending"
            Height          =   315
            Index           =   0
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   555
            Width           =   1095
         End
         Begin VB.CheckBox chkSortD 
            Caption         =   "Ascending"
            Height          =   315
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtSortD 
            Height          =   315
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "SPAREPART TYPE"
            ToolTipText     =   "Default Field"
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox cboFldSort 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   360
            TabIndex        =   37
            Top             =   1185
            Width           =   2175
         End
         Begin VB.ComboBox cboFldSort 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   360
            TabIndex        =   36
            Top             =   870
            Width           =   2175
         End
         Begin VB.ComboBox cboFldSort 
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   35
            Top             =   555
            Width           =   2175
         End
         Begin VB.Label lblNoD 
            Caption         =   "1"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   270
            Width           =   375
         End
         Begin VB.Label lblNo 
            Caption         =   "4"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   1215
            Width           =   375
         End
         Begin VB.Label lblNo 
            Caption         =   "3"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   900
            Width           =   375
         End
         Begin VB.Label lblNo 
            Caption         =   "2"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   585
            Width           =   375
         End
      End
      Begin lvButton.lvButtons_H cmdSortOK 
         Height          =   375
         Left            =   840
         TabIndex        =   30
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "SPbckOrdSum.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdSortCancel 
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
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
         Image           =   "SPbckOrdSum.frx":05B6
         cBack           =   -2147483633
      End
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Top             =   8760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin TabDlg.SSTab tabSQD 
      Height          =   2415
      Left            =   3240
      TabIndex        =   16
      Top             =   8520
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Sales Qty. Date Covered"
      TabPicture(0)   =   "SPbckOrdSum.frx":0B50
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   3735
         Begin MSComCtl2.DTPicker dTo 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Format          =   87556097
            CurrentDate     =   39500
         End
         Begin MSComCtl2.DTPicker dFrom 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Format          =   87556097
            CurrentDate     =   39500
         End
         Begin VB.Label Label2 
            Caption         =   "Date To:"
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Date From:"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
      End
      Begin lvButton.lvButtons_H cmdOk 
         Height          =   375
         Left            =   840
         TabIndex        =   22
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "SPbckOrdSum.frx":0B6C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
         Image           =   "SPbckOrdSum.frx":1106
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   7800
      Width           =   11415
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   180
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblSalesCov 
         Alignment       =   1  'Right Justify
         Caption         =   "SALES QTY. NOT YET UPDATED"
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
         Left            =   4800
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   195
         Width           =   6495
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
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   6615
      Left            =   0
      TabIndex        =   6
      Tag             =   "XY"
      Top             =   1200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11668
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin CRVIEWERLibCtl.CRViewer vwrNew 
      Height          =   7335
      Left            =   0
      TabIndex        =   5
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
      EnableExportButton=   -1  'True
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
         TabIndex        =   7
         Top             =   285
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8040
         TabIndex        =   8
         Top             =   120
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
         Image           =   "SPbckOrdSum.frx":16A0
         cBack           =   -2147483633
      End
      Begin VB.Label lblRemarks 
         Caption         =   "All Items"
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
         Top             =   240
         Width           =   7335
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
      Image           =   "SPbckOrdSum.frx":1C3A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   1
      Top             =   0
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
      Image           =   "SPbckOrdSum.frx":21D4
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
      Image           =   "SPbckOrdSum.frx":276E
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
      Image           =   "SPbckOrdSum.frx":2D08
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   12
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Update Sales"
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
      Image           =   "SPbckOrdSum.frx":32A2
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   13
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Refresh"
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
      Image           =   "SPbckOrdSum.frx":383C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   11760
      TabIndex        =   15
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
      Image           =   "SPbckOrdSum.frx":3DD6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   38
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Sort Records"
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
      Image           =   "SPbckOrdSum.frx":4370
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11805
      m_MDI_Height    =   8925
   End
   Begin VB.Menu nmuSales 
      Caption         =   "Sales Report"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "SPbckOrdSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset
Dim rsPRV As New ADODB.Recordset
Dim BtnIndx As Integer, Qry As String, OrderFld As String
Dim fRow As Integer, fCol As Integer
Public ActRow As Integer


Private Sub cboFldSort_Click(Index As Integer)
    If Index = 2 Then Exit Sub
    lblNo(Index + 1).ForeColor = QBColor(1)
    cboFldSort(Index + 1).Enabled = True
    chkSort(Index + 1).Enabled = True
End Sub

Private Sub cboFldSort_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSortOK.SetFocus
End Sub

Private Sub cboFldSort_LostFocus(Index As Integer)
    Dim i As Integer
    
    With cboFldSort(Index)
        If .ListIndex <= 0 Then
            cboFldSort(Index).Text = ""
            chkSort(Index).Caption = "Ascending"
            For i = Index + 1 To 2
                lblNo(i).ForeColor = QBColor(8)
                cboFldSort(i).Text = ""
                cboFldSort(i).Enabled = False
                chkSort(i).Enabled = False
                chkSort(i).Value = 0
                chkSort(i).Caption = "Ascending"
            Next i
        End If
    End With
End Sub

Private Sub chkSort_Click(Index As Integer)
    If chkSort(Index).Value = 1 Then
        chkSort(Index).Caption = "Descending"
    Else
        chkSort(Index).Caption = "Ascending"
    End If
End Sub

Private Sub chkSortD_Click()
    If chkSortD.Value = 1 Then
        chkSortD.Caption = "Descending"
    Else
        chkSortD.Caption = "Ascending"
    End If
End Sub

Private Sub cmdCancel_Click()
    tabSQD.Visible = False
    For i = 2 To 7
        cmdTrans(i).Enabled = True
    Next i
    flxSP.Enabled = True
    Frame1.Enabled = True
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSP
        For i = fRow To .Rows - 1
            For j = fCol To 6
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 6:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 6, i + 1, i)
                    fCol = IIf(j = 6, 1, j + 1)
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
Public Sub ShowFilteredRecords()
    Dim SelFrom As String
    
    ' SelFrom = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.REMARKS, SPAREPARTS.SPAREPARTID, SPAREPARTS.SUPPLIERCODE From SPAREPARTS"
    'Qry = SelFrom & " " & txtCriteria.Text & " " & IIf(txtOrder.Text <> "", txtOrder.Text & ";", "ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.PARTNUMBER;")
 
    If txtOrder.Text <> "" Then
        txOrder = txtOrder.Text & ";"
    Else
        txOrder = "ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.PARTNUMBER;"
    End If
    
    SelFrom = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.ORDEREF, SPAREPARTS.REMARKS, SPAREPARTS.SPAREPARTID, SPAREPARTS.SUPPLIERCODE From SPAREPARTS"
    'Qry = SelFrom & " " & txtCriteria.Text & " " & txOrder
    Qry = SelFrom & " " & txtCriteria.Text & " " & IIf(txtOrder.Text <> "", txtOrder.Text & ";", "ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.PARTNUMBER;")
 
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
    Call RXResizer1_ResizeComplete
    With flxSP
        If .Rows > 1 Then
            .Col = 1: .Row = 1
            .ColSel = 8:  .RowSel = 1
            ActRow = 1
            lblRec.Caption = "RECORD " & ActRow & " OF " & flxSP.Rows - 1
        End If
    End With
End Sub


Private Sub RefreshSP()
    Dim i As Integer, Ttl As Currency
    Dim dF As String, dT As String, myRem As String, SelFrom As String
    
    SelFrom = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.ORDEREF, SPAREPARTS.REMARKS, SPAREPARTS.SPAREPARTID, SPAREPARTS.SUPPLIERCODE From SPAREPARTS"
    If txtOrder.Text = "" Then
        Qry = SelFrom & " ORDER BY SPAREPARTS.REMARKS, SPAREPARTS.SUPPLIERCODE, SPAREPARTS.PARTNUMBER;"
    Else
        Qry = SelFrom & " " & txtOrder.Text & ";"
    End If
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
        If .Rows > 1 Then
            .Col = 1: .Row = 1
            .ColSel = 8:  .RowSel = 1
            ActRow = 1
            lblRec.Caption = "RECORD " & ActRow & " OF " & flxSP.Rows - 1
        End If
    End With
    Call RXResizer1_ResizeComplete
    dF = GetSQDsettings("SALES QTY FROM", 1)
    dT = GetSQDsettings("SALES QTY TO", 1)
    If Not IsDate(dF) And Not IsDate(dT) Then
        lblSalesCov.Caption = "SALES QTY. NOT YET UPDATED"
        dFrom.Value = DateValue("12/19/2007")
        dTo.Value = Date
    ElseIf Not IsDate(dF) Then
        lblSalesCov.Caption = "SALES QTY. NOT YET UPDATED [Date From]"
        dFrom.Value = DateValue("12/19/2007")
    ElseIf Not IsDate(dT) Then
        lblSalesCov.Caption = "SALES QTY. NOT YET UPDATED [Date To]"
        dTo.Value = Date
    ElseIf IsDate(dF) And IsDate(dT) Then
        dFrom.Value = DateValue(dF)
        dTo.Value = DateValue(dT)
        myRem = GetSQDsettings("SALES QTY FROM", 2)
        lblSalesCov.Caption = "Sales Qty. updated from " & dFrom.Value & " to " & dTo.Value & IIf(myRem <> "", " [" & myRem & "]", "")
    End If
    lblRemarks.Caption = "You are currently viewing all items"
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, sQty As Integer
    Dim dCov As String, myRem As String
    
    'Validate if there are records to update
    If flxSP.Rows = 2 And flxSP.TextMatrix(1, 2) = "" Then
        MsgBox "No records available to update!", vbExclamation + vbOKOnly, "Update Sales Qty."
        Call cmdCancel_Click
        Exit Sub
    End If
    tabSQD.Visible = False
    lblRec.Visible = False
    lblSalesCov.Visible = False
    pBar.Visible = True
    pBar.Value = 0
    With flxSP
        pBar.Max = .Rows - 1
        For i = 1 To .Rows - 1
            sQty = GetSalesQty(.TextMatrix(i, 7))
            .TextMatrix(i, 5) = sQty
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.SALES = " & sQty & " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 7) & "));"
            
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.SALES = " & sQty & " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 7) & ";"
            
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
    End With
    'Update Settings
    myRem = IIf(InStr(UCase$(lblRemarks.Caption), "ALL ITEMS") > 0, "All Items", "Selected Items")
    'LYRcon.Execute "UPDATE SETTINGS SET SETTINGS.[VALUE] = '" & dFrom.Value & "', SETTINGS.REMARKS = '" & myRem & "' WHERE (((SETTINGS.DESCRIPTION)='SALES QTY FROM'));"
    'LYRcon.Execute "UPDATE SETTINGS SET SETTINGS.[VALUE] = '" & dTo.Value & "', SETTINGS.REMARKS = '" & myRem & "' WHERE (((SETTINGS.DESCRIPTION)='SALES QTY TO'));"
    'mySQL
    LYRcon.Execute "UPDATE SETTINGS SET SETTINGS.VALUE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "', SETTINGS.REMARKS = '" & myRem & "' WHERE SETTINGS.DESCRIPTION='SALES QTY FROM';"
    LYRcon.Execute "UPDATE SETTINGS SET SETTINGS.VALUE = '" & Format(dTo.Value, "yyyy-mm-dd") & "', SETTINGS.REMARKS = '" & myRem & "' WHERE SETTINGS.DESCRIPTION='SALES QTY TO';"
    
    pBar.Visible = False
    lblRec.Visible = True
    lblSalesCov.Visible = True
    dCov = "Sales Qty. updated from " & dFrom.Value & " to " & dTo.Value & "[" & myRem & "]"
    lblSalesCov.Caption = dCov
    MsgBox "Sales Qty. successfully updated!" & vbCrLf & dCov, vbInformation + vbOKOnly, "Sales Qty."
    Call cmdCancel_Click
End Sub

Function GetSalesQty(MyID As Long) As Integer
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & dFrom.Value & "# And (SPITR.TRANSDATE)<=#" & dTo.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT Sum(SPITR.`OUT`) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(dTo.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.SPAREPARTID=" & MyID & ";"
    
    With rsQty
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetSalesQty = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            GetSalesQty = 0
        End If
        .Close
    End With
    
End Function

Private Sub cmdSortCancel_Click()
    TabSort.Visible = False
    For i = 2 To 7
        cmdTrans(i).Enabled = True
    Next i
    flxSP.Enabled = True
    Frame1.Enabled = True

End Sub

Private Sub cmdSortOK_Click()
    Dim Arrfld As Variant

    Arrfld = Split(OrderFld, "*")
    
    txtOrder.Text = "ORDER BY SPAREPARTS.REMARKS" & IIf(chkSortD.Value = 1, " DESC", "") & ", SPAREPARTS.SUPPLIERCODE"
    If cboFldSort(0).ListIndex <= 0 Then
        txtOrder.Text = txtOrder.Text & ", SPAREPARTS.PARTNUMBER"
    Else
        txtOrder.Text = txtOrder.Text & ", " & Arrfld(cboFldSort(0).ListIndex) & IIf(chkSort(0).Value = 1, " DESC", "")
    End If
    If cboFldSort(1).ListIndex > 0 Then txtOrder.Text = txtOrder.Text & ", " & Arrfld(cboFldSort(1).ListIndex) & IIf(chkSort(1).Value = 1, " DESC", "")
    If cboFldSort(2).ListIndex > 0 Then txtOrder.Text = txtOrder.Text & ", " & Arrfld(cboFldSort(2).ListIndex) & IIf(chkSort(2).Value = 1, " DESC", "")
    If txtCriteria.Text = "" Then
        Call RefreshSP
    Else
        Call ShowFilteredRecords
    End If
    Call cmdSortCancel_Click
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            'Call RefreshSP
            Call PrintPreView(False)
        Case 1  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
        Case 2  'Print Preview
            BtnIndx = 7
            tmrFloatRev.Enabled = True
            Call DisplayReport
            Call PrintPreView(True)
        Case 3  'Filter
            For i = 0 To 2
                 With SPfilter.cboLookIn(i)
                   .AddItem "PART NUMBER"
                   .AddItem "DESCRIPTION"
                   .AddItem "QTY.-ON-HAND"
                   .AddItem "BACK ORDER"
                   .AddItem "P.O. REFERENCE"
                   .AddItem "SPARE PARTS TYPE"
                   .AddItem "SUPPLIER"
                 End With
            Next i
            SPfilter.fldNames = "SPAREPARTS.PARTNUMBER*SPAREPARTS.DESCRIPTION*SPAREPARTS.QTYONHAND*SPAREPARTS.BACKORDER*SPAREPARTS.ORDEREF*SPAREPARTS.REMARKS*SPAREPARTS.SUPPLIERCODE"
            SPfilter.fldNdx = "23"
            SPfilter.fldDates = ""
            SPfilter.fldSupplier = 6
            SPfilter.MySource = "INVBCKSAL"
            SPfilter.Show vbModal
        Case 4  'Sort records
            For i = 2 To 7
                cmdTrans(i).Enabled = False
            Next i
            flxSP.Enabled = False
            Frame1.Enabled = False
            With TabSort
                .Top = (Me.Height - (.Height + 1000)) / 2
                .Left = (Me.Width - .Width) / 2
                .Visible = True
            End With
        Case 5  'Update Sales Qty.
            For i = 2 To 7
                cmdTrans(i).Enabled = False
            Next i
            flxSP.Enabled = False
            Frame1.Enabled = False
            With tabSQD
                .Top = (Me.Height - (.Height + 1000)) / 2
                .Left = (Me.Width - .Width) / 2
                .Visible = True
            End With
        Case 6  'Refresh
            Call RefreshSP
        Case 7  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
    End Select
End Sub

Function GetSQDsettings(MyDesc As String, myFld As Integer) As String
    Dim rsSetting As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SETTINGS.* From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='" & MyDesc & "'));"
    
    'mySQL
    Qry = "SELECT SETTINGS.* From SETTINGS WHERE SETTINGS.DESCRIPTION='" & MyDesc & "';"
    
    With rsSetting
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetSQDsettings = IIf(IsNull(.Fields(myFld)), "", .Fields(myFld))
        Else
            GetSQDsettings = ""
        End If
        .Close
    End With

End Function

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
    'dtpFrom.Value = Date
    Call RefreshSP
    Call PrintPreView(False)
    
    OrderFld = "NONE*SPAREPARTS.PARTNUMBER*SPAREPARTS.DESCRIPTION*SPAREPARTS.QTYONHAND*SPAREPARTS.BACKORDER*SPAREPARTS.ORDEREF"
    For i = 0 To 2
        cboFldSort(i).AddItem "(None)"
        cboFldSort(i).AddItem "PART NUMBER"
        cboFldSort(i).AddItem "DESCRIPTION"
        cboFldSort(i).AddItem "QTY-ON-HAND"
        cboFldSort(i).AddItem "BACK ORDER"
        cboFldSort(i).AddItem "P.O. REFERENCE"
        If i > 0 Then lblNo(i).ForeColor = QBColor(8)
    Next i
    MotoRep.sBar.Panels(6).Text = "Inventory Summary w/ Back Orders and Sales "
End Sub

Private Sub DisplayReport()
    Dim crIBS As New rptInvBckSal
    Dim MyUser As String
    
    MyUser = GetUserName
    
    If rsPRV.State = adStateOpen Then rsPRV.Close
    rsPRV.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    crIBS.DiscardSavedData
    crIBS.Database.SetDataSource rsPRV
    crIBS.EnableParameterPrompting = False
    crIBS.ParameterFields(1).ClearCurrentValueAndRange
    crIBS.ParameterFields(1).AddCurrentValue "*Note: " & lblSalesCov.Caption
    crIBS.ParameterFields(2).ClearCurrentValueAndRange
    crIBS.ParameterFields(2).AddCurrentValue "Prepared By: " & MyUser
    crIBS.ParameterFields(3).ClearCurrentValueAndRange
    crIBS.ParameterFields(3).AddCurrentValue SysBranch
    crIBS.PaperSource = crPRBinTractor
    crIBS.PaperSize = crDefaultPaperSize
    vwrNew.ReportSource = crIBS
    vwrNew.ViewReport
End Sub

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
        Frame1.Visible = False
        Frame2.Visible = False
        flxSP.Visible = False
        vwrNew.Visible = True
    Else
        vwrNew.Visible = False
        Frame1.Visible = True
        Frame2.Visible = True
        flxSP.Visible = True
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
End Sub
Private Sub RXResizer1_ResizeComplete()
    
    With flxSP
        .FormatString = "   |<PART NUMBER |<DESCRIPTION|^QTY-ON-HAND|^BACK ORDER |^P.O. REF.|TYPE    ||^CODE"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.18    'Partno
        .ColWidth(2) = (.Width - 300) * 0.3   'Desc
        .ColWidth(3) = (.Width - 300) * 0.1    'Qty
        .ColWidth(4) = (.Width - 300) * 0.1    'BO
        .ColWidth(5) = (.Width - 300) * 0.18    'Sales
        .ColWidth(6) = (.Width - 300) * 0.12    'Remarks
        .ColWidth(7) = 0                        'SPID
        .ColWidth(8) = (.Width - 300) * 0.05    'Supplier Code
    End With
    'cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = lblRemarks.Top
    txtFind.Top = lblRemarks.Top
    txtFind.Left = Frame1.Width - txtFind.Width - 50
    cmdFind.Left = txtFind.Left - (cmdFind.Width + 50)
End Sub
Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 8 Then
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
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 2 'Index of Button view grid
            For i = 2 To 7
                cmdTrans(i).Top = -375
                cmdTrans(i).Left = cmdTrans(2).Width * (i - 2)
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

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub
