VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form SPTopSales 
   Caption         =   "Top Selling Spare Parts"
   ClientHeight    =   9480
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   4440
      TabIndex        =   52
      Top             =   9600
      Visible         =   0   'False
      Width           =   6015
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   1095
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   5775
         DisplayGroupTree=   -1  'True
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
   End
   Begin TabDlg.SSTab TabSort 
      Height          =   2535
      Left            =   10920
      TabIndex        =   40
      Top             =   9960
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPTopSales.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOK"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CboSort1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CboAD1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CboSort2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CboSort3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CboAD2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CboAD3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.ComboBox CboAD3 
         Height          =   315
         Left            =   4320
         TabIndex        =   49
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox CboAD2 
         Height          =   315
         Left            =   4320
         TabIndex        =   48
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox CboSort3 
         Height          =   315
         Left            =   1320
         TabIndex        =   47
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox CboSort2 
         Height          =   315
         Left            =   1320
         TabIndex        =   46
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox CboAD1 
         Height          =   315
         Left            =   4320
         TabIndex        =   43
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox CboSort1 
         Height          =   315
         Left            =   1320
         TabIndex        =   42
         Top             =   360
         Width           =   2895
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   375
         Left            =   2040
         TabIndex        =   50
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "OK"
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
         Image           =   "SPTopSales.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   3720
         TabIndex        =   51
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Cancel"
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
         Image           =   "SPTopSales.frx":05B6
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "3. SORT BY"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   45
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "2. SORT BY"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "1. SORT BY"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   41
         Top             =   390
         Width           =   975
      End
   End
   Begin VB.TextBox txtRemarks 
      Height          =   285
      Left            =   5760
      TabIndex        =   39
      Top             =   9600
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtTopSalesID 
      Height          =   285
      Left            =   4560
      TabIndex        =   38
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   9600
      TabIndex        =   30
      Tag             =   "xy"
      Top             =   840
      Width           =   4695
      Begin VB.TextBox txtNote 
         Height          =   285
         Left            =   840
         MaxLength       =   100
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   240
         Width           =   2700
      End
      Begin lvButton.lvButtons_H cmdFilter 
         Height          =   375
         Left            =   840
         TabIndex        =   31
         Tag             =   "xy"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   661
         Caption         =   "FILTER"
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
         Image           =   "SPTopSales.frx":0B50
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdSort 
         Height          =   375
         Left            =   2760
         TabIndex        =   32
         Tag             =   "xy"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   661
         Caption         =   "SORT ORDER"
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
         Image           =   "SPTopSales.frx":10EA
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdStart 
         Height          =   375
         Left            =   840
         TabIndex        =   33
         Tag             =   "xy"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   661
         Caption         =   "START"
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
         Image           =   "SPTopSales.frx":1684
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         Caption         =   "NOTE:"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Tag             =   "xy"
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "DATE:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Tag             =   "xy"
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "TYPE OF ITEMS"
      Height          =   1575
      Left            =   5160
      TabIndex        =   25
      Tag             =   "XY"
      Top             =   840
      Width           =   4335
      Begin VB.ComboBox cboSPType 
         Height          =   315
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtTop 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1800
         TabIndex        =   29
         Tag             =   "XYF"
         Text            =   "100"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox ChkTop 
         Caption         =   "SAVE ONLY TOP"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin MSComCtl2.UpDown UpdTop 
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Tag             =   "XY"
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   100
         Increment       =   10
         Max             =   5000
         Min             =   10
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTop 
         Caption         =   "FROM ALL PARTS"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   10560
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Tag             =   "XYF"
      Top             =   9150
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
      Left            =   5880
      Top             =   360
   End
   Begin VB.TextBox txtYr 
      Height          =   285
      Left            =   7800
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   6960
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6360
      Top             =   360
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   6615
      Left            =   120
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   2520
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11668
      _Version        =   393216
      RowHeightMin    =   285
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "SALES PERIOD"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   840
      Width           =   4935
      Begin MSComCtl2.UpDown UpDShow 
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   3
         Max             =   12
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "SHOW LAST 3"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin MSComCtl2.UpDown UpdPrev 
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Tag             =   "xy"
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         Max             =   12
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.OptionButton OptPrev 
         Caption         =   "FOR THE LAST 1 "
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   720
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptBeg 
         Caption         =   "FROM BEGINNING"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblShow 
         Caption         =   "MONTHS SALES QTY. IN REPORT"
         Height          =   200
         Left            =   1920
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   1227
         Width           =   2775
      End
      Begin VB.Label lblPrev 
         Caption         =   "MONTH"
         Height          =   200
         Left            =   2880
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   747
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   0
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
      Image           =   "SPTopSales.frx":1C1E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   6
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Su&mmary"
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
      Image           =   "SPTopSales.frx":21B8
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   9720
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Date   |<PO Number    |<Orders"
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&New Summary"
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
      Image           =   "SPTopSales.frx":2752
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   15
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit "
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
      Image           =   "SPTopSales.frx":2CEC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Save "
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
      Image           =   "SPTopSales.frx":3286
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPTopSales.frx":3820
      cBack           =   -2147483633
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
      TabIndex        =   12
      Tag             =   "XY"
      Top             =   9150
      Width           =   615
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14550
      m_MDI_Height    =   9990
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOP SELLING SPARE PARTS"
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
      Left            =   9720
      TabIndex        =   2
      Top             =   450
      Width           =   4500
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
      Top             =   450
      Width           =   2615
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
      Top             =   475
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOP SELLING SPARE PARTS"
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
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuINV 
      Caption         =   "SP Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuProcess 
         Caption         =   "Process Inventory"
      End
   End
End
Attribute VB_Name = "SPTopSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSP As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Dim M_Proj As CRAXDRT.Application
Dim M_Report As CRAXDRT.Report
Dim ComboSrch As Boolean

Private Sub BtnSaveCancel(Bln As Boolean)
    If Bln Then
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = False
        cmdTrans(2).Visible = False
        cmdTrans(3).Visible = False
        cmdTrans(4).Left = 0
        cmdTrans(4).Visible = True
        cmdTrans(5).Left = cmdTrans(4).Width
        cmdTrans(5).Visible = True
    Else
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = True
        cmdTrans(2).Visible = True
        cmdTrans(3).Visible = True
    End If
End Sub


Private Sub cboSPtype_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If ComboSrch Then
        ComboSrch = False
        Exit Sub
    End If
    With cboSPType
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                ComboSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With

End Sub

Private Sub cboSPtype_Click()
    Call cboSPtype_Change
End Sub

Private Sub cboSPtype_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSPType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSPtype_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then ComboSrch = True
End Sub

Private Sub cboSPtype_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub ChkShow_Click()
    'IF UpDShow.Value = 0 THE
    'Call UpDShow_Change
    If ChkShow.Value = 0 Then UpDShow.Value = 0
    
End Sub

Private Sub ChkTop_Click()
    ChkTop.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Dim i As Integer
    
    For i = 0 To 3
        cmdTrans(i).Enabled = True
    Next i
    Frame3.Enabled = True
    flxSP.Enabled = True
    TabSort.Visible = False
    TabSort.Top = 13000
    CboSort1.ListIndex = Val(CboSort1.Tag)
    CboSort2.ListIndex = Val(CboSort2.Tag)
    CboSort3.ListIndex = Val(CboSort3.Tag)
    CboAD1.ListIndex = Val(CboAD1.Tag)
    CboAD2.ListIndex = Val(CboAD2.Tag)
    CboAD3.ListIndex = Val(CboAD3.Tag)

End Sub

Private Sub cmdFilter_Click()
    Dim i As Integer
    
    For i = 0 To 2
         With SPfilter.cboLookIn(i)
           .AddItem "PART NUMBER"
           .AddItem "DESCRIPTION"
           .AddItem "QTY-ON-HAND"
           .AddItem "BACK ORDER"
           .AddItem "AVG. MONTHLY"
           .AddItem "TOTAL SALES"
           .AddItem "COST"
           .AddItem "RANK"
         End With
    Next i
    SPfilter.fldNames = "SPAREPARTS.PARTNUMBER*SPAREPARTS.DESCRIPTION*TOPSALEDTLS.QTYONHAND*TOPSALEDTLS.BACKORDER*TOPSALEDTLS.AVGMOSALES*TOPSALEDTLS.TTLSALESQTY*TOPSALEDTLS.COST*TOPSALEDTLS.RANK"
    SPfilter.fldNdx = "234567" 'start w/ 0
    SPfilter.fldDates = ""
    SPfilter.fldSupplier = 10   'if not included in filter use number > last index of field
    SPfilter.MySource = "SPTOPSALES"
    SPfilter.Show vbModal

End Sub

Private Sub cmdOK_Click()
    Dim Qry As String, MySort As String, myRem As String
    Dim i As Integer
    
    'Remove Order statement in Qry
    i = InStr(txtSQL.Text, "ORDER BY")
    Qry = Left$(txtSQL.Text, i - 1)
    'Add Sort Order to Query
    If CboSort1.ListIndex = -1 Then
        MySort = "TOPSALEDTLS.TTLSALESQTY DESC"
    Else
        If CboSort1.ListIndex = 0 Then
            MySort = "SPAREPARTS.PARTNUMBER " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 1 Then
            MySort = "SPAREPARTS.DESCRIPTION " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 2 Then
            MySort = "TOPSALEDTLS.QTYONHAND " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 3 Then
            MySort = "TOPSALEDTLS.BACKORDER " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 4 Then
            MySort = "TOPSALEDTLS.AVGMOSALES " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 5 Then
            MySort = "TOPSALEDTLS.TTLSALESQTY " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 6 Then
            MySort = "TOPSALEDTLS.COST " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 7 Then
            MySort = "TOPSALEDTLS.RANK " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        End If
        myRem = "[SORTED BY: " & CboSort1.Text & IIf(CboAD1.ListIndex <> -1, "->" & CboAD1.Text, "->ASC") & "]"
        If CboSort2.ListIndex <> -1 Then
            If CboSort2.ListIndex = 0 Then
                MySort = MySort & ", SPAREPARTS.PARTNUMBER " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 1 Then
                MySort = MySort & ", SPAREPARTS.DESCRIPTION " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 2 Then
                MySort = MySort & ", TOPSALEDTLS.QTYONHAND " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 3 Then
                MySort = MySort & ", TOPSALEDTLS.BACKORDER " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 4 Then
                MySort = MySort & ", TOPSALEDTLS.AVGMOSALES " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 5 Then
                MySort = MySort & ", TOPSALEDTLS.TTLSALESQTY " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 6 Then
                MySort = MySort & ", TOPSALEDTLS.COST " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 7 Then
                MySort = MySort & ", TOPSALEDTLS.RANK " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            End If
            myRem = "[SORTED BY: " & CboSort1.Text & IIf(CboAD1.ListIndex <> -1, "->" & CboAD1.Text, "->ASC") & ", " & CboSort2.Text & IIf(CboAD2.ListIndex <> -1, "->" & CboAD2.Text, "->ASC") & "]"
            If CboSort3.ListIndex <> -1 Then
                If CboSort3.ListIndex = 0 Then
                    MySort = MySort & ", SPAREPARTS.PARTNUMBER " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 1 Then
                    MySort = MySort & ", SPAREPARTS.DESCRIPTION " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 2 Then
                    MySort = MySort & ", TOPSALEDTLS.QTYONHAND " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 3 Then
                    MySort = MySort & ", TOPSALEDTLS.BACKORDER " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 4 Then
                    MySort = MySort & ", TOPSALEDTLS.AVGMOSALES " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 5 Then
                    MySort = MySort & ", TOPSALEDTLS.TTLSALESQTY " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 6 Then
                    MySort = MySort & ", TOPSALEDTLS.COST " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 7 Then
                    MySort = MySort & ", TOPSALEDTLS.RANK " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                End If
                myRem = "[SORTED BY: " & CboSort1.Text & IIf(CboAD1.ListIndex <> -1, "->" & CboAD1.Text, "->ASC") & ", " & CboSort2.Text & IIf(CboAD2.ListIndex <> -1, "->" & CboAD2.Text, "->ASC") & ", " & CboSort3.Text & IIf(CboAD3.ListIndex <> -1, "->" & CboAD3.Text, "->ASC") & "]"
            End If
        End If
    End If
    Qry = Qry & " ORDER BY " & MySort & ";"
    txtRemarks.Text = txtRemarks.Text & " " & myRem
    'Assign Complete Query back to txtSQL
    txtSQL.Text = Qry
    flxSP.Clear
    If rsSP.State = adStateOpen Then rsSP.Close
    rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSP.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 11
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsSP
    'Arrange FlxSP
    flxSP.ColWidth(0) = 300
    flxSP.ColWidth(1) = 0
    flxSP.ColWidth(2) = (flxSP.Width - 300) * 0.14
    flxSP.ColWidth(3) = (flxSP.Width - 300) * 0.22
    flxSP.ColWidth(flxSP.Cols - 1) = 0
    For i = 4 To flxSP.Cols - 1
        flxSP.ColWidth(i) = (flxSP.Width - 300) * 0.08
    Next i
    flxSP.TextMatrix(0, 4) = "QH"
    flxSP.TextMatrix(0, 5) = "BO"
    If UpDShow.Value > 0 Then
        Ans = UpDShow.Value
        For i = 6 To 5 + UpDShow.Value
            flxSP.TextMatrix(0, i) = UCase$(Format$(DateAdd("m", Ans * -1, Date), "MMM. 'YY"))
            Ans = Ans - 1
        Next i
    End If
    flxSP.TextMatrix(0, flxSP.Cols - 4) = "AVG MO."
    flxSP.TextMatrix(0, flxSP.Cols - 3) = "TTL. SALES"
    flxSP.TextMatrix(0, flxSP.Cols - 2) = "COST"
    flxSP.TextMatrix(0, flxSP.Cols - 1) = "RANK"
    'Format Records
    pBar.Value = 0
    pBar.Max = flxSP.Rows - 1
    pBar.Visible = True
    With flxSP
        For i = 1 To .Rows - 1
            .TextMatrix(i, .Cols - 4) = Format(.TextMatrix(i, .Cols - 4), "#,##0.00")
            .TextMatrix(i, .Cols - 2) = Format(.TextMatrix(i, .Cols - 2), "#,##0.00")
            pBar.Value = pBar.Value + 1
            pPct = (pBar.Value / pBar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
            DoEvents
        Next i
    End With
    pBar.Visible = False
    lblTotal.Caption = ""
    For i = 0 To 3
        cmdTrans(i).Enabled = True
    Next i
    Frame3.Enabled = True
    flxSP.Enabled = True
    TabSort.Visible = False
    TabSort.Top = 13000

End Sub

Private Sub cmdSort_Click()
    Dim i As Integer
    
    For i = 0 To 3
        cmdTrans(i).Enabled = False
    Next i
    Frame3.Enabled = False
    flxSP.Enabled = False
    With TabSort
        .Top = (Me.Height - .Height) \ 2
        .Left = (Me.Width - .Width) \ 2
        .Caption = ""
        .Visible = True
        .ZOrder
    End With
    CboSort1.Tag = CboSort1.ListIndex
    CboSort2.Tag = CboSort2.ListIndex
    CboSort3.Tag = CboSort3.ListIndex
    CboAD1.Tag = CboAD1.ListIndex
    CboAD2.Tag = CboAD2.ListIndex
    CboAD3.Tag = CboAD3.ListIndex
    
End Sub

Private Sub cmdStart_Click()
    Dim rsTopS As New ADODB.Recordset
    Dim rsTopD As New ADODB.Recordset
    Dim rsEx As New ADODB.Recordset
    Dim Qry As String, MyCov As String, MyItems As String
    Dim Ans As Integer, i As Integer, j As Integer, pPct As Single, SumSales As Integer
    Dim DCov1 As Date, DCov2 As Date
    
    '.... Validation ....
    If cboSPType.ListIndex = -1 Then
        MsgBox "Please select Type of Spare Part from List!", vbCritical + vbOKOnly, "Validation"
        Exit Sub
    End If
    If OptBeg.Value = True Then
        MyCov = "From Beginning"
    Else
        MyCov = Format$(DateAdd("m", Val(UpdPrev.Value) * -1, Date), "MMM. yyyy") & " to " & Format$(DateAdd("M", -1, Date), "MMM. yyyy")
    End If
    MyItems = cboSPType.Text
    Ans = MsgBox("Do you want to process " & MyCov & " for " & MyItems & "?", vbQuestion + vbYesNo, MyCov)
    If Ans = vbNo Then Exit Sub
    cmdTrans(0).Tag = ""
    '***Access
    'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.REMARKS)='" & cboSPType.Text & "')) ORDER BY SPAREPARTS.QTYONHAND DESC;"
    'mySQL
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER " & _
          "From SPAREPARTS WHERE SPAREPARTS.REMARKS='" & cboSPType.Text & "' ORDER BY SPAREPARTS.QTYONHAND DESC;"
    
    If rsSP.State = adStateOpen Then rsSP.Close
    rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSP.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 11
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsSP
    flxSP.Refresh
    flxSP.Cols = 10 + UpDShow.Value
    flxSP.ColWidth(0) = 300
    flxSP.ColWidth(1) = 0
    flxSP.ColWidth(2) = (flxSP.Width - 300) * 0.14
    flxSP.ColWidth(3) = (flxSP.Width - 300) * 0.22
    flxSP.ColWidth(flxSP.Cols - 1) = 0
    For i = 4 To flxSP.Cols - 1
        flxSP.ColWidth(i) = (flxSP.Width - 300) * 0.08
    Next i
    flxSP.TextMatrix(0, 4) = "QH"
    flxSP.TextMatrix(0, 5) = "BO"
    If UpDShow.Value > 0 Then
        Ans = UpDShow.Value
        For i = 6 To 5 + UpDShow.Value
            flxSP.TextMatrix(0, i) = UCase$(Format$(DateAdd("m", Ans * -1, Date), "MMM. 'YY"))
            Ans = Ans - 1
        Next i
    End If
    flxSP.TextMatrix(0, flxSP.Cols - 4) = "AVG. MO"
    flxSP.TextMatrix(0, flxSP.Cols - 3) = "TTL. SALES"
    flxSP.TextMatrix(0, flxSP.Cols - 2) = "COST"
    flxSP.TextMatrix(0, flxSP.Cols - 1) = "RANK"
    'disable objects
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    cmdTrans(4).Enabled = False
    cmdTrans(5).Enabled = False
    'Start Progress Bar
    pBar.Value = 0
    pBar.Max = flxSP.Rows - 1
    pBar.Visible = True
    'Change Here
    With flxSP
        For i = .Rows - 1 To 1 Step -1
            'Start Progress bar
            If OptBeg.Value = True Then
                .TextMatrix(i, .Cols - 3) = GetSalesQty(Val(.TextMatrix(i, 1)), DateValue("12/18/2007"), DateSerial(Year(Date), Month(Date), 1))
            Else
                If UpDShow.Value > 0 Then
                    DCov1 = DateAdd("m", UpDShow.Value * -1, DateSerial(Year(Date), Month(Date), 1))
                    DCov2 = DateSerial(Year(Date), Month(Date), 1)
                    .TextMatrix(i, .Cols - 3) = GetSalesQty(Val(.TextMatrix(i, 1)), Format(DCov1, "yyyy-mm-dd"), Format(DCov2, "yyyy-mm-dd"))
                End If
            End If
            If Val(.TextMatrix(i, .Cols - 3)) = 0 Then .RemoveItem i   'Do not include items without sales
            pBar.Value = pBar.Value + 1
            DoEvents
            pPct = (pBar.Value / pBar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
            DoEvents
        Next i
        DoEvents
    End With
    Ans = MsgBox("The next process requires you to SAVE the transaction" & vbCrLf & "to get the top " & txtTop.Text & " items." & vbCrLf & vbCrLf & _
          "Do you want to continue?", vbQuestion + vbYesNo, "SAVE")
    If Ans = vbNo Then
        cmdTrans(4).Enabled = True
        cmdTrans(5).Enabled = True
        Call cmdTrans_Click(5)
        Exit Sub
    End If
    'Saving Records
    '***Access
    'With rsTopS
    '    .Open "TOPSALESUMM", LYRcon, adOpenKeyset, adLockOptimistic
    '    .AddNew
    '    !SalesPeriod = IIf(OptBeg.Value = True, 0, UpdPrev.Value)
    '    !MonthInRpt = UpDShow.Value
    '    !TopValue = Val(txtTop.Text)
    '    !ItemType = cboSPType.Text
    '    txtNote.Text = Replace$(txtNote.Text, "'", "`")
    '    !NOTE = txtNote.Text
    '    !DateCreated = txtDate.Text
    '    .Update
    '    txtTopSalesID.Text = !TopSalesID
    '    .Close
    'End With
    'mySQL
    'If rsEx.State = adStateOpen Then rsEx.Close
    'rsEx.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    'If rsSP.RecordCount > 0 Then
    '     LYRcon.Execute "Delete * from topsalesumm where topsalesumm.DATECREATED = '" & txtDate.Text & "';"
    '    rsEx.Close
    'End If
    
    txtNote.Text = Replace$(txtNote.Text, "'", "`")
    dtCreated = Format(txtDate.Text, "yyyy-mm-dd hh:mm:ss")
    
    LYRcon.Execute "Insert Into topsalesumm(SALESPERIOD, MONTHINRPT, TOPVALUE, ITEMTYPE, NOTE, DATECREATED)" & _
        "VALUES(" & IIf(OptBeg.Value = True, 0, UpdPrev.Value) & ", " & UpDShow.Value & ", " & Val(txtTop.Text) & ", " & _
        "'" & cboSPType.Text & "', '" & txtNote.Text & "', '" & dtCreated & "');"
    
    Qry = "SELECT topsalesumm.TOPSALESID, topsalesumm.DateCreated From topsalesumm Where topsalesumm.DATECREATED = '" & Format(dtCreated, "yyyy-mm-dd hh:mm:ss") & "';"
    txtTopSalesID.Text = GetFieldValue(Qry)
    
    pBar.Value = 0
    pBar.Max = flxSP.Rows - 1
    pBar.Visible = True
    
    '***Access
    'With rsTopD
    '    .Open "TOPSALEDTLS", LYRcon, adOpenKeyset, adLockOptimistic
    '    For i = 1 To flxSP.Rows - 1
    '        .AddNew
    '        !SparepartID = Val(flxSP.TextMatrix(i, 1))
    '        !QtyOnHand = Val(flxSP.TextMatrix(i, 4))
    '        !BACKORDER = Val(flxSP.TextMatrix(i, 5))
    '        !TTLSALESQTY = Val(flxSP.TextMatrix(i, flxSP.Cols - 3))
    '        !TopSalesID = Val(txtTopSalesID.Text)
    '        .Update
    '        Pbar.Value = Pbar.Value + 1
    '        pPct = (Pbar.Value / Pbar.Max)
    '        lblTotal.Caption = Format$(pPct, "##0.0%")
    '        DoEvents
    '    Next i
    '    .Close
    'End With
    
    'mySQL
    For i = 1 To flxSP.Rows - 1
            
        LYRcon.Execute "Insert Into topsaledtls(SPAREPARTID, QTYONHAND, BACKORDER, TTLSALESQTY, TOPSALESID)" & _
            "VALUES(" & Val(flxSP.TextMatrix(i, 1)) & ", " & Val(flxSP.TextMatrix(i, 4)) & ", " & Val(flxSP.TextMatrix(i, 5)) & ", " & _
            "" & Val(flxSP.TextMatrix(i, flxSP.Cols - 3)) & ", " & Val(txtTopSalesID.Text) & ");"
        
        pBar.Value = pBar.Value + 1
        pPct = (pBar.Value / pBar.Max)
        lblTotal.Caption = Format$(pPct, "##0.0%")
        DoEvents
        
    Next i
    '-----------------------------
    'Build Qry to Display in FlxSP
    If UpDShow.Value > 0 Then
        MyCov = ""
        For i = 1 To UpDShow.Value
            MyCov = MyCov & " TOPSALEDTLS.MONTH" & i & ","
        Next i
        '***Access
        'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, " & MyCov & " TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, SPAREPARTS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((TOPSALEDTLS.TOPSALESID) = " & Val(txtTopSalesID.Text) & ")) ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        'mySQL
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, " & MyCov & " TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, SPAREPARTS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where TOPSALEDTLS.TOPSALESID = " & Val(txtTopSalesID.Text) & " ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        
    Else
        '***Access
        'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, SPAREPARTS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((TOPSALEDTLS.TOPSALESID) = " & Val(txtTopSalesID.Text) & ")) ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        'mySQL
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, SPAREPARTS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where TOPSALEDTLS.TOPSALESID = " & Val(txtTopSalesID.Text) & " ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
    
    End If
    flxSP.Clear
    If rsSP.State = adStateOpen Then rsSP.Close
    rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSP.RecordCount > 0 Then
        If Val(txtTop.Text) < rsSP.RecordCount Then
            rsSP.Move Val(txtTop.Text)
            j = rsSP!TTLSALESQTY
        Else
            rsSP.MoveLast
            j = rsSP!TTLSALESQTY
        End If
    Else
        j = 0
    End If
    rsSP.Close
    
    '***Access
    'LYRcon.Execute "DELETE TOPSALEDTLS.* From TOPSALEDTLS WHERE (((TOPSALEDTLS.TTLSALESQTY)<" & j & ") AND ((TOPSALEDTLS.TOPSALESID)=" & Val(txtTopSalesID.Text) & "));"
    'mySQL
    LYRcon.Execute "DELETE From TOPSALEDTLS WHERE TOPSALEDTLS.TTLSALESQTY<" & j & " AND TOPSALEDTLS.TOPSALESID=" & Val(txtTopSalesID.Text) & ";"
    
    'REopen File
    rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSP.DataSource = rsSP
    'Arrange FlxSP
    flxSP.ColWidth(0) = 300
    flxSP.ColWidth(1) = 0
    flxSP.ColWidth(2) = (flxSP.Width - 300) * 0.14
    flxSP.ColWidth(3) = (flxSP.Width - 300) * 0.22
    flxSP.ColWidth(flxSP.Cols - 1) = 0
    For i = 4 To flxSP.Cols - 1
        flxSP.ColWidth(i) = (flxSP.Width - 300) * 0.08
    Next i
    flxSP.TextMatrix(0, 4) = "QH"
    flxSP.TextMatrix(0, 5) = "BO"
    If UpDShow.Value > 0 Then
        Ans = UpDShow.Value
        For i = 6 To 5 + UpDShow.Value
            flxSP.TextMatrix(0, i) = UCase$(Format$(DateAdd("m", Ans * -1, Date), "MMM. 'YY"))
            Ans = Ans - 1
        Next i
    End If
    flxSP.TextMatrix(0, flxSP.Cols - 4) = "AVG MO."
    flxSP.TextMatrix(0, flxSP.Cols - 3) = "TTL. SALES"
    flxSP.TextMatrix(0, flxSP.Cols - 2) = "COST"
    flxSP.TextMatrix(0, flxSP.Cols - 1) = "RANK"
    'Get Monthly Sales
    If OptPrev.Value = True Then
        pBar.Value = 0
        pBar.Max = flxSP.Rows - 1
        pBar.Visible = True
        With flxSP
            For i = 1 To .Rows - 1
                Ans = UpDShow.Value
                SumSales = 0
                For j = 6 To 5 + UpDShow.Value  'Get Monthly Sales
                    DCov1 = DateAdd("m", Ans * -1, DateSerial(Year(Date), Month(Date), 1))
                    DCov2 = DateAdd("m", 1, DCov1)
                    .TextMatrix(i, j) = GetSalesQty(Val(.TextMatrix(i, 1)), DCov1, DCov2)
                    SumSales = SumSales + Val(.TextMatrix(i, j))
                    Ans = Ans - 1
                Next j
                .TextMatrix(i, .Cols - 4) = Format$(SumSales / UpDShow.Value, "###0.00")
                'Update Table TopSaleDtls
                .TextMatrix(i, .Cols - 2) = Format(.TextMatrix(i, .Cols - 2), "#,##0.00")
                .TextMatrix(i, .Cols - 1) = i
                MyCov = ""
                For j = 1 To UpDShow.Value
                    MyCov = MyCov & " TOPSALEDTLS.MONTH" & j & " = " & Val(.TextMatrix(i, 5 + j)) & ","
                Next j  'Change field to Dcov1
                '***Access
                'LYRcon.Execute "UPDATE TOPSALEDTLS SET " & MyCov & " TOPSALEDTLS.COST = " & Val(Replace(.TextMatrix(i, .Cols - 2), ",", "")) & ", TOPSALEDTLS.RANK = " & i & ", TOPSALEDTLS.AVGMOSALES = " & Val(.TextMatrix(i, .Cols - 4)) & " WHERE (((TOPSALEDTLS.TOPSALESID)=" & Val(txtTopSalesID.Text) & ") AND ((TOPSALEDTLS.SPAREPARTID)=" & Val(.TextMatrix(i, 1)) & "));"
                'mySQL
                LYRcon.Execute "UPDATE TOPSALEDTLS SET " & MyCov & " TOPSALEDTLS.COST = " & Val(Replace(.TextMatrix(i, .Cols - 2), ",", "")) & ", TOPSALEDTLS.RANK = " & i & ", TOPSALEDTLS.AVGMOSALES = " & Val(.TextMatrix(i, .Cols - 4)) & " WHERE TOPSALEDTLS.TOPSALESID=" & Val(txtTopSalesID.Text) & " AND TOPSALEDTLS.SPAREPARTID=" & Val(.TextMatrix(i, 1)) & ";"
                pBar.Value = pBar.Value + 1
                pPct = (pBar.Value / pBar.Max)
                lblTotal.Caption = Format$(pPct, "##0.0%")
                DoEvents
            Next i
        End With
    End If
    pBar.Visible = False
    lblTotal.Caption = ""
    MsgBox "Top " & txtTop.Text & " selling Spare Parts now ready!", vbInformation + vbOKOnly, "Finish"
    
    cmdTrans(4).Enabled = True
    cmdTrans(5).Enabled = True
    Frame3.Enabled = True
    txtDate.Locked = True
    txtNote.Locked = True
    cmdStart.Visible = False
    cmdFilter.Visible = True
    cmdSort.Visible = True
    Call BtnSaveCancel(False)

End Sub

Function GetSalesQty(MyID As Long, dFrom As Date, dTo As Date) As Long
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    '***Access
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & dFrom & "# And (SPITR.TRANSDATE)<#" & dTo & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & MyId & "));"
    
    'mySQL
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<'" & Format(dTo, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.SPAREPARTID=" & MyID & ";"
    
    MySales = Val(GetFieldValue(Qry))
    '***Access
    'Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyId & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & dFrom & "# And (SPITR.TRANSDATE)<#" & dTo & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(dFrom, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<'" & Format(dTo, "yyyy-mm-dd") & "';"
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

Private Sub cmdTrans_Click(Index As Integer)
    Dim dCov As String, Qry As String
    Dim Ans As Integer
    
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=97));"
            
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=97;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            Frame1.Enabled = True
            Frame2.Enabled = True
            Frame3.Enabled = True
            cmdStart.Visible = True
            cmdFilter.Visible = False
            cmdSort.Visible = False
            txtDate.Text = Format$(Now, "mm/dd/yyyy hh:mm:ss")
            Call BtnSaveCancel(True)
            cmdTrans(0).Tag = "NOT STARTED"
            txtRemarks.Text = ""
        Case 1  'Print
            If Val(txtTopSalesID.Text) = 0 Then
                MsgBox "No Records to Print!", vbExclamation + vbOKOnly, "Print"
                Exit Sub
            End If
            If cmdTrans(1).Tag = "" Then 'Show Print Preview
                cmdTrans(1).Caption = "Close Preview"
                cmdTrans(1).Tag = "Print Preview"
                Frame1.Visible = False
                Frame2.Visible = False
                Frame3.Visible = False
                flxSP.Visible = False
                Frame4.Top = Frame1.Top
                Frame4.Left = Frame1.Left
                Frame4.Width = flxSP.Width
                Frame4.Height = Frame1.Height + flxSP.Height
                CRViewer1.Top = 100
                CRViewer1.Left = 50
                CRViewer1.Width = Frame4.Width - 100
                CRViewer1.Height = Frame4.Height - 200
                Frame4.Visible = True
                Frame4.ZOrder
                Call PrintSPTopSales
            Else    'Close Print Preview
                cmdTrans(1).Caption = "&Print Preview"
                cmdTrans(1).Tag = ""
                Frame1.Visible = True
                Frame2.Visible = True
                Frame3.Visible = True
                flxSP.Visible = True
                Frame4.Top = 13000
                Frame4.Visible = False
            End If
            
        Case 2  'Summary
            FrmSummary.mySrc = "Top Selling Spare Parts"
            FrmSummary.Show
            Unload Me
        Case 3  'Exit
            Unload Me
            MotoRep.tbMain.Visible = True
        Case 4  'Save
            If cmdTrans(0).Tag = "NOT STARTED" Then
                MsgBox "Transaction has not yet started!" & vbCrLf & "Click the START button to finalize transaction", vbInformation + vbOKOnly, "Save"
            End If
        Case 5  'Cancel
            Ans = MsgBox("Do you want to Cancel Summary?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            Frame1.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            cmdStart.Visible = False
            cmdFilter.Visible = False
            cmdSort.Visible = False
            Call BtnSaveCancel(False)
            flxSP.Clear
            flxSP.Rows = 2
            flxSP.Cols = 11
            Call RXResizer1_ResizeComplete
    End Select
End Sub

Private Sub PrintSPTopSales()
    Dim txtH As TextObject
    Dim dbfldObj As FieldObject
    Dim lnObj As LineObject
    Dim txtF As TextObject
    Dim PageNumber As FieldObject
    Dim myRem As String
    Dim i As Integer, MaxC As Integer, Cw As Integer, Lw As Integer, myLeft As Integer
    
    Set M_Proj = New CRAXDRT.Application
    Set M_Report = M_Proj.NewReport
    
    M_Report.DiscardSavedData
    M_Report.Database.AddADOCommand rsSP.ActiveConnection, rsSP.ActiveCommand
    
    'Set papersize and orientation
    If UpDShow.Value > 6 Then
        M_Report.PaperSize = crPaperLegal
        MaxC = 17000
    Else
        M_Report.PaperSize = crPaperLetter
        MaxC = 15000
    End If
    M_Report.PaperOrientation = crLandscape
    
    ' Add Section 1 : Report header
    M_Report.Sections(1).Height = 800
    Set txtH = M_Report.Sections(1).AddTextObject(SysBranch, 10, 0)
    txtH.Name = "Title"
    With txtH
        .Width = MaxC: .Height = 300: .Font.Name = "ARIAL": .Font.Size = 12: .Font.Bold = True: .Top = 50: .Left = 200: .HorAlignment = crHorCenterAlign
    End With
    Set txtH = M_Report.Sections(1).AddTextObject("LYR BLDG. II, Sobrecary St., Tagum City", 10, 0)
    With txtH
        .Width = MaxC: .Height = 300: .Font.Name = "ARIAL": .Font.Size = 10: .Font.Bold = True: .Top = 300: .Left = 200: .HorAlignment = crHorCenterAlign
    End With

    'Add Section 2: Page Header
    M_Report.Sections(2).Height = 900
    myRem = IIf(cboSPType.ListIndex >= 1, cboSPType.Text, "SPARE") & " PARTS SALES PERFORMANCE FROM " & IIf(OptBeg.Value = True, "BEGINNING", UCase$(Format$(DateAdd("M", Val(UpdPrev.Value) * -1, txtDate.Text), "MMM. 'YY")) & " TO " & UCase$(Format$(DateAdd("M", -1, txtDate.Text), "MMM. 'YY")))
    Set txtH = M_Report.Sections(2).AddTextObject(myRem, 10, 10)
    With txtH
        .Width = MaxC: .Height = 300: .Font.Size = 10: .Font.Name = "ARIAL": .Top = 0: .Left = 200: .HorAlignment = crLeftAlign
    End With
    If txtRemarks.Text <> "" Then
        Set txtH = M_Report.Sections(2).AddTextObject(txtRemarks.Text, 10, 10)
        With txtH
            .Width = MaxC: .Height = 300: .Font.Size = 10: .Font.Name = "ARIAL": .Top = 300: .Left = 200: .HorAlignment = crLeftAlign
        End With
    End If
    Set lnObj = M_Report.Sections(2).AddLineObject(200, 600, MaxC, 600) 'Horizontal Line top of page Header
    lnObj.LineThickness = 1
    Set lnObj = M_Report.Sections(2).AddLineObject(200, 600, 200, 900) 'Vertical Line Left of Part Number
    lnObj.LineThickness = 1
    Set txtH = M_Report.Sections(2).AddTextObject("PART NUMBER", 10, 10)
    With txtH
        .Width = 2000: .Height = 300: .Font.Size = 10: .Font.Name = "ARIAL": .Top = 620: .Left = 200: .HorAlignment = crHorCenterAlign
    End With
    Set lnObj = M_Report.Sections(2).AddLineObject(2000, 600, 2000, 900) 'Vertical Line Left of Description
    lnObj.LineThickness = 1
    Set txtH = M_Report.Sections(2).AddTextObject("DESCRIPTION", 2010, 10)
    With txtH
        .Width = 3000: .Height = 300: .Font.Size = 10: .Font.Name = "ARIAL": .Top = 620: .Left = 2000: .HorAlignment = crHorCenterAlign
    End With
    Set lnObj = M_Report.Sections(2).AddLineObject(5000, 600, 5000, 900) 'Vertical Line Left of QH
    lnObj.LineThickness = 1
    Cw = (MaxC - 5000) \ (flxSP.Cols - 4)
    myLeft = 5000
    flxSP.TextMatrix(0, flxSP.Cols - 3) = "T-SALES"
    For i = 4 To flxSP.Cols - 1
        Set txtH = M_Report.Sections(2).AddTextObject(flxSP.TextMatrix(0, i), myLeft, 10)
        With txtH
            .Width = Cw: .Height = 300: .Font.Size = 8: .Font.Name = "ARIAL": .Top = 620: .Left = myLeft: .HorAlignment = crHorCenterAlign
        End With
        myLeft = myLeft + Cw
        Set lnObj = M_Report.Sections(2).AddLineObject(myLeft, 600, myLeft, 900) 'Vertical Line Left of Current Display Field
        lnObj.LineThickness = 1
    Next i
    Set lnObj = M_Report.Sections(2).AddLineObject(MaxC, 600, MaxC, 900) 'Vertical Line last column
    lnObj.LineThickness = 1
    Set lnObj = M_Report.Sections(2).AddLineObject(200, 900, MaxC, 900)
    lnObj.LineThickness = 1
    
    'Add Section 3: Details section
    M_Report.Sections(3).Height = 300
    Set lnObj = M_Report.Sections(3).AddLineObject(200, 0, 200, 300) 'Vertical Line Left of Part Number
    lnObj.LineThickness = 1
    Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.PartNumber}", 50, 0)
    With dbfldObj
        .Width = 1800: .Height = 260: .Font.Size = 8: .Font.Name = "TAHOMA": .Top = 0: .Left = 300: .HorAlignment = crLeftAlign
    End With
    Set lnObj = M_Report.Sections(3).AddLineObject(2000, 0, 2000, 300) 'Vertical Line Left of Description
    lnObj.LineThickness = 1
    Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.Description}", 50, 0)
    With dbfldObj
        .Width = 2900: .Height = 260: .Font.Size = 8: .Font.Name = "TAHOMA": .Top = 0: .Left = 2100: .HorAlignment = crLeftAlign
    End With
    Cw = (MaxC - 5000) \ (flxSP.Cols - 4)
    myLeft = 5000
    Set lnObj = M_Report.Sections(3).AddLineObject(5000, 0, 5000, 300) 'Vertical Line Left of QH
    lnObj.LineThickness = 1
    Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.QTYONHAND}", 50, 0)
    With dbfldObj
        .Width = Cw: .Height = 260: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 0: .Left = myLeft: .HorAlignment = crHorCenterAlign
    End With
    myLeft = myLeft + Cw
    Set lnObj = M_Report.Sections(3).AddLineObject(myLeft, 0, myLeft, 300) 'Vertical Line Left of BO
    lnObj.LineThickness = 1
    Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.BACKORDER}", 50, 0)
    With dbfldObj
        .Width = Cw: .Height = 260: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 0: .Left = myLeft: .HorAlignment = crHorCenterAlign
    End With
    myLeft = myLeft + Cw
    Set lnObj = M_Report.Sections(3).AddLineObject(myLeft, 0, myLeft, 300) 'Vertical Line Left of NEXT FIELD
    lnObj.LineThickness = 1
    If UpDShow.Value > 0 Then
        For i = 1 To UpDShow.Value
            Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.Month" & i & "}", 50, 0)
            With dbfldObj
                .Width = Cw: .Height = 260: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 0: .Left = myLeft: .HorAlignment = crHorCenterAlign
            End With
            myLeft = myLeft + Cw
            Set lnObj = M_Report.Sections(3).AddLineObject(myLeft, 0, myLeft, 300) 'Vertical Line Left of NEXT FIELD
            lnObj.LineThickness = 1
        Next i
    End If
    Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.AVGMOSALES}", 50, 0)
    With dbfldObj
        .Width = Cw: .Height = 260: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 0: .Left = myLeft: .HorAlignment = crHorCenterAlign: .DecimalPlaces = 2
    End With
    myLeft = myLeft + Cw
    Set lnObj = M_Report.Sections(3).AddLineObject(myLeft, 0, myLeft, 300) 'Vertical Line Left of TOTAL SALES
    lnObj.LineThickness = 1
    Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.TTLSALESQTY}", 50, 0)
    With dbfldObj
        .Width = Cw: .Height = 260: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 0: .Left = myLeft: .HorAlignment = crHorCenterAlign
    End With
    myLeft = myLeft + Cw
    Set lnObj = M_Report.Sections(3).AddLineObject(myLeft, 0, myLeft, 300) 'Vertical Line Left of COST
    lnObj.LineThickness = 1
    Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.COST}", 50, 0)
    With dbfldObj
        .Width = Cw: .Height = 260: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 0: .Left = myLeft: .HorAlignment = crHorCenterAlign: .DecimalPlaces = 2
    End With
    myLeft = myLeft + Cw
    Set lnObj = M_Report.Sections(3).AddLineObject(myLeft, 0, myLeft, 300) 'Vertical Line Left of RANK
    lnObj.LineThickness = 1
    Set dbfldObj = M_Report.Sections(3).AddFieldObject("{ado.RANK}", 50, 0)
    With dbfldObj
        .Width = Cw: .Height = 260: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 0: .Left = myLeft: .HorAlignment = crHorCenterAlign
    End With
    Set lnObj = M_Report.Sections(3).AddLineObject(MaxC, 0, MaxC, 300) 'Vertical Line LAST COLUMN
    lnObj.LineThickness = 1
    Set lnObj = M_Report.Sections(3).AddLineObject(200, 300, MaxC, 300)
    lnObj.LineThickness = 1
    
    'Add Section 5: page Footer
    M_Report.Sections(5).Height = 500
    Set txtH = M_Report.Sections(5).AddTextObject("Date Printed: " & Format$(Now, "mm/dd/yyyy hh:ss ampm"), 10, 10)
    With txtH
        .Width = 4000: .Height = 300: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 50: .Left = 200: .HorAlignment = crLeftAlign
    End With
    Set PageNumber = M_Report.Sections(5).AddSpecialVarFieldObject(crSVTPageNofM, MaxC - (3000), 50)
    With PageNumber
        .Width = 3000: .Height = 300: .Font.Size = 9: .Font.Name = "TAHOMA": .Top = 50: .Left = MaxC - 3000: .HorAlignment = crRightAlign
    End With
    'Show Report
    CRViewer1.ReportSource = M_Report
    CRViewer1.DisplayGroupTree = False
    CRViewer1.ViewReport
    CRViewer1.Zoom 100
End Sub

Private Sub flxSP_EnterCell()
    ActRow = flxSP.Row
End Sub

Private Sub Form_Load()
    Dim rsSPType As New ADODB.Recordset
    Dim Qry As String
    
    MotoRep.sBar.Panels(6).Text = "Summary of Top Selling Spare Parts"
    ActRow = 1
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    CboSort1.AddItem "PART NUMBER"
    CboSort1.AddItem "DESCRIPTION"
    CboSort1.AddItem "QTY-ON-HAND"
    CboSort1.AddItem "BACK ORDER"
    CboSort1.AddItem "AVG. MONTHLY"
    CboSort1.AddItem "TOTAL SALES"
    CboSort1.AddItem "COST"
    CboSort1.AddItem "RANK"
    CboSort2.AddItem "PART NUMBER"
    CboSort2.AddItem "DESCRIPTION"
    CboSort2.AddItem "QTY-ON-HAND"
    CboSort2.AddItem "BACK ORDER"
    CboSort2.AddItem "AVG. MONTHLY"
    CboSort2.AddItem "TOTAL SALES"
    CboSort2.AddItem "COST"
    CboSort2.AddItem "RANK"
    CboSort3.AddItem "PART NUMBER"
    CboSort3.AddItem "DESCRIPTION"
    CboSort3.AddItem "QTY-ON-HAND"
    CboSort3.AddItem "BACK ORDER"
    CboSort3.AddItem "AVG. MONTHLY"
    CboSort3.AddItem "TOTAL SALES"
    CboSort3.AddItem "COST"
    CboSort3.AddItem "RANK"
    CboAD1.AddItem "ASC"
    CboAD1.AddItem "DESC"
    CboAD2.AddItem "ASC"
    CboAD2.AddItem "DESC"
    CboAD3.AddItem "ASC"
    CboAD3.AddItem "DESC"
    CboSort1.ListIndex = 7
    CboAD1.ListIndex = 0
    
    Qry = "SELECT DISTINCT SUPPLIER.SPTYPE From SUPPLIER ORDER BY SUPPLIER.SPTYPE;"
    With rsSPType
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            cboSPType.Clear
            cboSPType.AddItem "ALL TYPE"
            .MoveFirst
            Do While Not .EOF
                cboSPType.AddItem !SPtype
                .MoveNext
            Loop
        End If
        .Close
    End With
    cboSPType.ListIndex = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsSP.State = adStateOpen Then rsSP.Close
End Sub



Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    lblHeading(2).Left = Frame3.Left + (Frame3.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame3.Left + (Frame3.Width - lblHeading(3).Width)
    
    With flxSP
        If Val(txtTopSalesID.Text) = 0 Then
                            '  1       2             3         4   5            6                                                            7                                                       8                                                9       10      11    12
            .FormatString = "   |  |< PART NUMBER  |<DESCRIPTION|^QH|^BO|^" & UCase$(Format(DateAdd("M", -3, Date), "MMM. 'YY")) & "|^" & UCase$(Format(DateAdd("M", -2, Date), "MMM. 'YY")) & "|^" & UCase$(Format(DateAdd("M", -1, Date), "MMM. 'YY")) & "|^AVG. MO|^TTL SALES|^COST|^RANK"
            .ColWidth(0) = 300
            .ColWidth(1) = 0
            For i = 2 To 12
                .ColWidth(i) = (.Width - 300) * 0.06
            Next i
            .ColWidth(2) = (.Width - 300) * 0.14
            .ColWidth(3) = (.Width - 300) * 0.22
            .ZOrder
        End If
    End With
    
   '
   'txtFind.Top = cmdFind.Top
End Sub


Public Sub ShowRecord(MyID As Integer)
    Dim rsSal As New ADODB.Recordset
    Dim Qry As String, MyCov As String, i As Integer, Ans As Integer
    
    '***Access
    'Qry = "SELECT TOPSALESUMM.* From TOPSALESUMM WHERE (((TOPSALESUMM.TOPSALESID)=" & MyId & "));"
    
    'mySQL
    Qry = "SELECT TOPSALESUMM.* From TOPSALESUMM WHERE TOPSALESUMM.TOPSALESID=" & MyID & ";"
    
    With rsSal
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtDate.Text = Format$(!DateCreated, "mm/dd/yyyy")
            If !SalesPeriod = 0 Then
                OptBeg.Value = True
                OptPrev.Value = False
            Else
                OptPrev.Value = True
                OptBeg.Value = False
                UpdPrev.Value = !SalesPeriod
                'Call UpdPrev_Change
            End If
            If !MonthInRpt = 0 Then
                ChkShow.Value = 0
            Else
                ChkShow.Value = 1
                UpDShow.Value = !MonthInRpt
                'Call UpDShow_Change
            End If
            cboSPType.Text = !ItemType
            txtTop.Text = !TopValue
            txtNote.Text = IIf(IsNull(!NOTE), "", !NOTE)
            txtTopSalesID.Text = !TopSalesID
        End If
        .Close
    End With
    'Build Qry to Display in FlxSP
    If UpDShow.Value > 0 Then
        MyCov = ""
        For i = 1 To UpDShow.Value
            MyCov = MyCov & " TOPSALEDTLS.MONTH" & i & ","
        Next i
        '***Access
        'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, " & MyCov & " TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((TOPSALEDTLS.TOPSALESID) = " & Val(txtTopSalesID.Text) & ")) ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        'mySQL
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, " & MyCov & " TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where TOPSALEDTLS.TOPSALESID = " & Val(txtTopSalesID.Text) & " ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        
        
    Else
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((TOPSALEDTLS.TOPSALESID) = " & Val(txtTopSalesID.Text) & ")) ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        'mySQL
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where TOPSALEDTLS.TOPSALESID = " & Val(txtTopSalesID.Text) & " ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        
    End If
    txtSQL.Text = Qry
    CboSort1.ListIndex = 7
    CboAD1.ListIndex = 0
    flxSP.Clear
    If rsSP.State = adStateOpen Then rsSP.Close
    rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSP.DataSource = rsSP
    'Arrange FlxSP
    flxSP.ColWidth(0) = 300
    flxSP.ColWidth(1) = 0
    flxSP.ColWidth(2) = (flxSP.Width - 300) * 0.14
    flxSP.ColWidth(3) = (flxSP.Width - 300) * 0.22
    flxSP.ColWidth(flxSP.Cols - 1) = 0
    For i = 4 To flxSP.Cols - 1
        flxSP.ColWidth(i) = (flxSP.Width - 300) * 0.08
    Next i
    flxSP.TextMatrix(0, 4) = "QH"
    flxSP.TextMatrix(0, 5) = "BO"
    If UpDShow.Value > 0 Then
        Ans = UpDShow.Value
        For i = 6 To 5 + UpDShow.Value
            flxSP.TextMatrix(0, i) = UCase$(Format$(DateAdd("m", Ans * -1, Date), "MMM. 'YY"))
            Ans = Ans - 1
        Next i
    End If
    flxSP.TextMatrix(0, flxSP.Cols - 4) = "AVG MO."
    flxSP.TextMatrix(0, flxSP.Cols - 3) = "TTL. SALES"
    flxSP.TextMatrix(0, flxSP.Cols - 2) = "COST"
    flxSP.TextMatrix(0, flxSP.Cols - 1) = "RANK"
    'Format Records
    pBar.Value = 0
    pBar.Max = flxSP.Rows - 1
    pBar.Visible = True
    With flxSP
        For i = 1 To .Rows - 1
            .TextMatrix(i, .Cols - 4) = Format(.TextMatrix(i, .Cols - 4), "#,##0.00")
            .TextMatrix(i, .Cols - 2) = Format(.TextMatrix(i, .Cols - 2), "#,##0.00")
            pBar.Value = pBar.Value + 1
            pPct = (pBar.Value / pBar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
            DoEvents
        Next i
    End With
    pBar.Visible = False
    lblTotal.Caption = ""
    Frame3.Enabled = True
    txtDate.Locked = True
    txtNote.Locked = True
    cmdFilter.Visible = True
    cmdSort.Visible = True
    txtRemarks.Text = ""
    
End Sub

Public Sub ShowFilteredRecords()
    Dim Qry As String, MyCov As String, MySort As String
    Dim i As Integer
    
    'Add Criteria in default Select
    If UpDShow.Value > 0 Then
        MyCov = ""
        For i = 1 To UpDShow.Value
            MyCov = MyCov & " TOPSALEDTLS.MONTH" & i & ","
        Next i
        '***Access
        'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((TOPSALEDTLS.TOPSALESID) = " & Val(txtTopSalesID.Text) & ")) ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        'mySQL
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where TOPSALEDTLS.TOPSALESID = " & Val(txtTopSalesID.Text) & " ORDER BY TOPSALEDTLS.TTLSALESQTY DESC;"
        
    
    Else
        'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And ((TOPSALEDTLS.TOPSALESID) = " & Val(txtTopSalesID.Text) & "))"
        'mySQL
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TOPSALEDTLS.QTYONHAND, TOPSALEDTLS.BACKORDER, TOPSALEDTLS.AVGMOSALES, TOPSALEDTLS.TTLSALESQTY, TOPSALEDTLS.COST, TOPSALEDTLS.RANK " & _
              "FROM TOPSALEDTLS INNER JOIN SPAREPARTS ON TOPSALEDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID " & txtSQL.Text & " And TOPSALEDTLS.TOPSALESID = " & Val(txtTopSalesID.Text) & ";"
        
        
    End If
    'Add Sort Order to Query
    If CboSort1.ListIndex = -1 Then
        MySort = "TOPSALEDTLS.TTLSALESQTY DESC"
    Else
        If CboSort1.ListIndex = 0 Then
            MySort = "SPAREPARTS.PARTNUMBER " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 1 Then
            MySort = "SPAREPARTS.DESCRIPTION " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 2 Then
            MySort = "TOPSALEDTLS.QTYONHAND " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 3 Then
            MySort = "TOPSALEDTLS.BACKORDER " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 4 Then
            MySort = "TOPSALEDTLS.AVGMOSALES " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 5 Then
            MySort = "TOPSALEDTLS.TTLSALESQTY " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 6 Then
            MySort = "TOPSALEDTLS.COST " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        ElseIf CboSort1.ListIndex = 7 Then
            MySort = "TOPSALEDTLS.RANK " & IIf(CboAD1.ListIndex <> -1, CboAD1.Text, "ASC")
        End If
        If CboSort2.ListIndex <> -1 Then
            If CboSort2.ListIndex = 0 Then
                MySort = MySort & ", SPAREPARTS.PARTNUMBER " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 1 Then
                MySort = MySort & ", SPAREPARTS.DESCRIPTION " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 2 Then
                MySort = MySort & ", TOPSALEDTLS.QTYONHAND " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 3 Then
                MySort = MySort & ", TOPSALEDTLS.BACKORDER " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 4 Then
                MySort = MySort & ", TOPSALEDTLS.AVGMOSALES " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 5 Then
                MySort = MySort & ", TOPSALEDTLS.TTLSALESQTY " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 6 Then
                MySort = MySort & ", TOPSALEDTLS.COST " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            ElseIf CboSort2.ListIndex = 7 Then
                MySort = MySort & ", TOPSALEDTLS.RANK " & IIf(CboAD2.ListIndex <> -1, CboAD2.Text, "ASC")
            End If
            If CboSort3.ListIndex <> -1 Then
                If CboSort3.ListIndex = 0 Then
                    MySort = MySort & ", SPAREPARTS.PARTNUMBER " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 1 Then
                    MySort = MySort & ", SPAREPARTS.DESCRIPTION " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 2 Then
                    MySort = MySort & ", TOPSALEDTLS.QTYONHAND " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 3 Then
                    MySort = MySort & ", TOPSALEDTLS.BACKORDER " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 4 Then
                    MySort = MySort & ", TOPSALEDTLS.AVGMOSALES " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 5 Then
                    MySort = MySort & ", TOPSALEDTLS.TTLSALESQTY " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 6 Then
                    MySort = MySort & ", TOPSALEDTLS.COST " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                ElseIf CboSort3.ListIndex = 7 Then
                    MySort = MySort & ", TOPSALEDTLS.RANK " & IIf(CboAD3.ListIndex <> -1, CboAD3.Text, "ASC")
                End If
            End If
        End If
    End If
    Qry = Qry & " ORDER BY " & MySort & ";"
    'Assign Complete Query back to txtSQL
    txtSQL.Text = Qry
    flxSP.Clear
    If rsSP.State = adStateOpen Then rsSP.Close
    rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSP.RecordCount = 0 Then
        flxSP.Clear
        flxSP.Rows = 2
        flxSP.Cols = 11
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSP.DataSource = rsSP
    'Arrange FlxSP
    flxSP.ColWidth(0) = 300
    flxSP.ColWidth(1) = 0
    flxSP.ColWidth(2) = (flxSP.Width - 300) * 0.14
    flxSP.ColWidth(3) = (flxSP.Width - 300) * 0.22
    flxSP.ColWidth(flxSP.Cols - 1) = 0
    For i = 4 To flxSP.Cols - 1
        flxSP.ColWidth(i) = (flxSP.Width - 300) * 0.08
    Next i
    flxSP.TextMatrix(0, 4) = "QH"
    flxSP.TextMatrix(0, 5) = "BO"
    If UpDShow.Value > 0 Then
        Ans = UpDShow.Value
        For i = 6 To 5 + UpDShow.Value
            flxSP.TextMatrix(0, i) = UCase$(Format$(DateAdd("m", Ans * -1, Date), "MMM. 'YY"))
            Ans = Ans - 1
        Next i
    End If
    flxSP.TextMatrix(0, flxSP.Cols - 4) = "AVG MO."
    flxSP.TextMatrix(0, flxSP.Cols - 3) = "TTL. SALES"
    flxSP.TextMatrix(0, flxSP.Cols - 2) = "COST"
    flxSP.TextMatrix(0, flxSP.Cols - 1) = "RANK"
    'Format Records
    pBar.Value = 0
    pBar.Max = flxSP.Rows - 1
    pBar.Visible = True
    With flxSP
        For i = 1 To .Rows - 1
            .TextMatrix(i, .Cols - 4) = Format(.TextMatrix(i, .Cols - 4), "#,##0.00")
            .TextMatrix(i, .Cols - 2) = Format(.TextMatrix(i, .Cols - 2), "#,##0.00")
            pBar.Value = pBar.Value + 1
            pPct = (pBar.Value / pBar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
            DoEvents
        Next i
    End With
    pBar.Visible = False
    lblTotal.Caption = ""
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub UpdPrev_Change()
    OptPrev.Caption = "FOR THE LAST " & UpdPrev.Value
    lblPrev.Caption = IIf(UpdPrev.Value > 1, "MONTHS", "MONTH")
    OptPrev.Value = True
End Sub

Private Sub UpDShow_Change()
    ChkShow.Caption = "SHOW LAST " & UpDShow.Value
    lblShow.Caption = IIf(UpDShow.Value > 1, "MONTHS", "MONTH") & " SALES QTY. IN REPORT"
    If UpDShow.Value = 0 Then
        ChkShow.Value = 0
    Else
        ChkShow.Value = 1
    End If
End Sub

Private Sub UpdTop_Change()
    txtTop.Text = UpdTop.Value
    
End Sub
