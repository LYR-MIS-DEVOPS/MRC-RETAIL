VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SPBackOrder 
   Caption         =   "Spare Parts Back Order Form"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   11565
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabWait 
      Height          =   1695
      Left            =   3480
      TabIndex        =   47
      Top             =   9360
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPbackOrder.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Label Label6 
         Caption         =   "Please wait... I'm processing you request..."
         Height          =   255
         Left            =   1200
         TabIndex        =   48
         Top             =   720
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   480
         Picture         =   "SPbackOrder.frx":001C
         Top             =   727
         Width           =   240
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000002&
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   320
         Left            =   45
         Top             =   45
         Width           =   4940
      End
   End
   Begin VB.TextBox txtOrderID 
      Height          =   285
      Left            =   10440
      TabIndex        =   46
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSuppID 
      Height          =   285
      Left            =   9600
      TabIndex        =   45
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   480
      TabIndex        =   36
      Tag             =   "XYF"
      Top             =   7580
      Width           =   6975
      Begin VB.TextBox txtBackOrder 
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
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   870
         Width           =   1215
      End
      Begin VB.TextBox txtReceived 
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
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   555
         Width           =   1215
      End
      Begin VB.TextBox txtOrder 
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
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblBckPct 
         Caption         =   "0.00% Back Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   885
         Width           =   2535
      End
      Begin VB.Label lblRecPct 
         Caption         =   "0.00% Served"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3600
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   585
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL BACK ORDER:"
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
         Index           =   2
         Left            =   360
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   885
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL QTY. RECEIVED:"
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
         Index           =   1
         Left            =   360
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   585
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL QTY. ORDERED:"
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
         Index           =   0
         Left            =   360
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   270
         Width           =   2535
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   555
      TabIndex        =   31
      Tag             =   "XY"
      Top             =   7275
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   4440
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtFind 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Part Number Exist @ row 5"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin TabDlg.SSTab tabFind 
      Height          =   1095
      Left            =   2400
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPbackOrder.frx":05A6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSearch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboMatch"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNext"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Find Next"
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   195
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   645
         Width           =   975
      End
      Begin VB.ComboBox cboMatch 
         Height          =   315
         ItemData        =   "SPbackOrder.frx":05C2
         Left            =   1080
         List            =   "SPbackOrder.frx":05C4
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   675
         Width           =   2295
      End
      Begin VB.ComboBox txtSearch 
         Height          =   315
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Find What:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Match:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   705
         Width           =   855
      End
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   300
      Left            =   10680
      TabIndex        =   21
      Tag             =   "XYF"
      ToolTipText     =   "Find Part Number"
      Top             =   1995
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
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
      ImgAlign        =   4
      Image           =   "SPbackOrder.frx":05C6
      cBack           =   -2147483633
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   10920
      Width           =   135
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   5295
      Left            =   480
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   2280
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^    |<PART NUMBER   |<DESCRIPTION  |^ORDER   |^RECEIVED|^BACK ORDER|<PACKING LIST  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   480
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   700
      Width           =   10695
      Begin VB.TextBox txtTerms 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   876
         Width           =   3735
      End
      Begin VB.ComboBox cboSupplier 
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
         Left            =   2040
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   192
         Width           =   3735
      End
      Begin VB.TextBox txtSaddress 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   538
         Width           =   3735
      End
      Begin VB.TextBox txtScode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   7800
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   876
         Width           =   2535
      End
      Begin VB.TextBox txtPONumber 
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
         Height          =   315
         Left            =   7800
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   538
         Width           =   2535
      End
      Begin VB.TextBox txtPODate 
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
         Height          =   315
         Left            =   7800
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   200
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "SUPPLIERADDRESS:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Tag             =   "XYF"
         Top             =   568
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "TERMS:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   906
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "SUPPLIER NAME:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER TYPE"
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
         Index           =   5
         Left            =   6600
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   906
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. NUMBER"
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
         Index           =   4
         Left            =   6240
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   568
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. DATE"
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
         Index           =   3
         Left            =   6600
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   230
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "SPbackOrder.frx":0B60
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   19
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Back Order"
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
      Image           =   "SPbackOrder.frx":10FA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   3360
      TabIndex        =   20
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Order"
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
      Image           =   "SPbackOrder.frx":1694
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   0
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11805
      m_MDI_Height    =   9720
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERED SPARE PARTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARE PARTS BACK ORDER FORM"
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
      Height          =   240
      Index           =   2
      Left            =   7590
      TabIndex        =   13
      Tag             =   "TRF"
      Top             =   450
      Width           =   3345
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
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Tag             =   "TLF"
      Top             =   450
      Width           =   2900
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
      Height          =   255
      Index           =   1
      Left            =   510
      TabIndex        =   12
      Tag             =   "TLF"
      Top             =   475
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARE PARTS BACK ORDER FORM"
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
      Height          =   240
      Index           =   3
      Left            =   7620
      TabIndex        =   14
      Tag             =   "TRF"
      Top             =   480
      Width           =   3345
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERED SPARE PARTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   2070
      Width           =   2415
   End
   Begin VB.Menu mnuRows 
      Caption         =   "SpareParts"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row        Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "SPBackOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSupplier As New ADODB.Recordset

Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean, ImportPO As Boolean
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim AddNew As Boolean, LockFlx As Boolean, Confirm As Boolean, SuppSrch As Boolean
Dim Qry As String

Private Sub ClearEntries()
    cboSupplier.Text = ""
    txtScode.Text = ""
    txtSaddress.Text = ""
    txtSuppID.Text = ""
    txtPODate.Text = ""
    txtPONumber.Text = ""
    cboType.Text = ""
    txtOrderID.Text = ""
    txtTerms.Text = ""
    flxParts.Rows = 1
    flxParts.Rows = 2
End Sub

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    LockFlx = bln
End Sub


Private Sub cboFind_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    With cboFind
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'If ActRow = i + 1 Then Exit Sub
            txtFind.Tag = i + 1 'IIf(LastRowLoaded >= i, i + 1, i)
            txtFind.Top = (flxParts.CellTop + flxParts.Top) - 315
            txtFind.Left = flxParts.CellLeft + flxParts.Left
            txtFind.Text = "Part Number Exist @ row " & Val(txtFind.Tag)
            txtFind.Visible = True
            txtFind.ZOrder
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                PrtSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtFind.Tag = ""
            txtFind.Visible = False
            txtFind.Text = ""
        End If
    End With
End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsSupplier.MoveFirst
            rsSupplier.Move i
            txtSuppID.Text = rsSupplier!SUPPLIERID
            txtScode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            txtSaddress.Text = rsSupplier!Address
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppID.Text = ""
            txtScode.Text = ""
            txtSaddress.Text = ""
        End If
    End With
End Sub

Private Sub cboSupplier_Click()
    Call cboSupplier_Change
End Sub

Private Sub cboSupplier_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtSaddress.SetFocus
End Sub

Private Sub cboSupplier_LostFocus()
    Dim rsBranch As New ADODB.Recordset
    
    If InStr(cboSupplier.Text, "HONDA PHIL") > 0 Then
        With rsBranch
            '***Access
            'Qry = "SELECT BRANCH.* From BRANCH WHERE (((BRANCH.BRANCHID)=1));"
            'mySQL
            Qry = "SELECT BRANCH.* From BRANCH WHERE BRANCH.BRANCHID=1;"
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                txtDealerCode.Text = IIf(IsNull(!MCDealerCode), "", !MCDealerCode)
                txtDealerName.Text = !branchName
                txtAddress.Text = !Address
            End If
            .Close
        End With
    Else
        txtDealerCode.Text = ""
        txtDealerName.Text = StrConv(SysBranch, vbProperCase)
        txtAddress.Text = StrConv(SysAddress, vbProperCase)
    End If
    If cboSupplier.Text <> "" Then
        txtPONumber.Text = GetOrderNo
    End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Call cboType_LostFocus
        flxParts.SetFocus
        flxParts.Col = 1
        flxParts.Row = flxParts.Rows - 1
    End If
End Sub

Private Sub cmdCancel_Click()
    tabFind.Visible = False
    cmdNext.Tag = 1
    cmdCancel.Tag = "hide"
    Call AddInCombo
    flxParts.SetFocus
End Sub

Private Sub AddInCombo()
    Dim X As Integer, w As Boolean
    
    If txtSearch.Text = "" Then Exit Sub
    w = False
    For X = 0 To txtSearch.ListCount - 1
        If txtSearch.Text = txtSearch.List(X) Then
            If X = 0 Then
                w = True
                Exit For
            Else
                txtSearch.RemoveItem X
                Exit For
            End If
        End If
    Next X
    If w = False Then txtSearch.AddItem txtSearch.Text, 0
End Sub

Private Sub cmdFind_Click()
    tabFind.Visible = True
    tabFind.Caption = ""
    If cboMatch.ListCount > 0 Then cboMatch.Text = cboMatch.List(1)
    cmdNext.Tag = 1
    cmdCancel.Tag = ""
    If Len(txtSearch.Text) > 0 Then
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
    End If
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    Dim i As Integer, j As Integer
    
    With flxParts
        For i = Val(cmdNext.Tag) To .Rows - 1
            For j = 1 To IIf(Confirm, 8, 5)
                If cboMatch.Text = "Whole Word" Then
                    If UCase$(txtSearch.Text) = UCase$(.TextMatrix(i, j)) Then
                        .SetFocus: .Row = i: .Col = 0: .ColSel = 8
                        .TopRow = i: cmdNext.Tag = i + 1
                        txtSearch.SelStart = 0: txtSearch.SelLength = Len(txtSearch.Text)
                        txtSearch.SetFocus
                        Call AddInCombo
                        Exit Sub
                    End If
                Else
                    If InStr(1, UCase$(.TextMatrix(i, j)), UCase$(txtSearch.Text)) > 0 Then
                        .SetFocus: .Row = i: .Col = 0: .ColSel = 5
                        .TopRow = i: cmdNext.Tag = i + 1
                        txtSearch.SelStart = 0: txtSearch.SelLength = Len(txtSearch.Text)
                        txtSearch.SetFocus
                        Call AddInCombo
                        Exit Sub
                    End If
                End If
            Next j
        Next i
        MsgBox "End of Search!" & vbCrLf & "No more records found!", vbInformation + vbOKOnly, "Find"
        cmdNext.Tag = 1
        Call AddInCombo
        .Col = 0: .Row = .Rows - 1
        .TopRow = .Rows - 1
        txtSearch.SetFocus
    End With
End Sub


Private Sub cmdTrans_Click(Index As Integer)

    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 21) As String
    Dim Qry As String
    
    Select Case Index
        Case 2  'Update Order
        Case 3  'View
            SPorderSum.Show
            Unload Me
        Case 4  'Print
            Ans = MsgBox("Do you want to Print Back Order?", vbQuestion + vbYesNo, "Back Order")
            If Ans = vbNo Then Exit Sub
            'Display tab
            With tabWait
                .TabCaption(0) = ""
                .Top = (Me.ScaleHeight - .Height) / 2
                .Left = (Me.ScaleWidth - .Width) / 2
                .Visible = True
                cmdTrans(3).Enabled = False
                cmdTrans(4).Enabled = False
                cmdTrans(5).Enabled = False
            End With
            '***Access
            'Qry = "SELECT BACKORDERDTLS.* From BACKORDERDTLS WHERE (((BACKORDERDTLS.SPORDERID)=" & Val(txtOrderID.Text) & "));"
            
            'mySQL
            Qry = "SELECT BACKORDERDTLS.* From BACKORDERDTLS WHERE BACKORDERDTLS.SPORDERID=" & Val(txtOrderID.Text) & ";"
            
            If CtrlNumExist(Qry) Then
                'Remove Previous records
                '***Access
                'LYRcon.Execute "DELETE BACKORDERDTLS.*, BACKORDERDTLS.SPORDERID From BACKORDERDTLS WHERE (((BACKORDERDTLS.SPORDERID)=" & Val(txtOrderID.Text) & "));"
                
                'mySQL
                LYRcon.Execute "DELETE From BACKORDERDTLS WHERE BACKORDERDTLS.SPORDERID=" & Val(txtOrderID.Text) & ";"
                
            End If
            With flxParts
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 5)) > 0 Then
                        LYRcon.Execute "INSERT INTO BACKORDERDTLS ( SPAREPARTID, ORDERQTY, RECEIVED, BACKORDER, PACKINGLIST, SPORDERID) " & _
                                       "VALUES (" & Val(.TextMatrix(i, 7)) & ", " & Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & Val(.TextMatrix(i, 5)) & ", '" & .TextMatrix(i, 6) & "', " & Val(txtOrderID.Text) & ");"
                    End If
                Next i
            End With
            With tabWait
                .TabCaption(0) = ""
                .Top = 20000
                .Visible = False
                cmdTrans(3).Enabled = True
                cmdTrans(4).Enabled = True
                cmdTrans(5).Enabled = True
            End With
            '***Access
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, BACKORDERDTLS.ORDERQTY, BACKORDERDTLS.RECEIVED, BACKORDERDTLS.BACKORDER, BACKORDERDTLS.PACKINGLIST " & _
                  "FROM BACKORDERDTLS INNER JOIN SPAREPARTS ON BACKORDERDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where (((BACKORDERDTLS.SPORDERID) = " & Val(txtOrderID.Text) & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
            
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, BACKORDERDTLS.ORDERQTY, BACKORDERDTLS.RECEIVED, BACKORDERDTLS.BACKORDER, BACKORDERDTLS.PACKINGLIST " & _
                  "FROM BACKORDERDTLS INNER JOIN SPAREPARTS ON BACKORDERDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID Where BACKORDERDTLS.SPORDERID = " & Val(txtOrderID.Text) & " ORDER BY SPAREPARTS.PARTNUMBER;"
            
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = cboSupplier.Text  'txtPONumber.Text
            PRNmc.ParamValue2 = txtSaddress.Text  'cboType.Text
            PRNmc.ParamValue3 = txtPONumber.Text
            PRNmc.ParamValue4 = txtPODate.Text
            PRNmc.ParamValue5 = Space$(5 - Len(txtOrder.Text)) & txtOrder.Text
            PRNmc.ParamValue6 = Space$(5 - Len(txtReceived.Text)) & txtReceived.Text & Space$(2) & lblRecPct.Caption
            PRNmc.ParamValue7 = Space$(5 - Len(txtBackOrder.Text)) & txtBackOrder.Text & Space$(2) & lblBckPct.Caption
            PRNmc.MySource = "BACKORDERPO"
            PRNmc.Show
        Case 5  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
    End Select

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    cmdA.Top = 15000
    Call LockData(True)
    Call LoadSpareParts
    cboType.AddItem "(01) - URGENT"
    cboType.AddItem "(04) - STOCK"
    cboMatch.AddItem "Whole Word"
    cboMatch.AddItem "Any Part of Word"
    
    MotoRep.sBar.Panels(6).Text = "Spare Parts Back Order Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub

Private Sub Label3_Click()
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch.Text)
    txtSearch.SetFocus
End Sub


Private Sub RXResizer1_ResizeComplete()
    With flxParts
        .ColWidth(0) = 1200
        .ColWidth(1) = (.Width - 1000) * 0.18
        .ColWidth(2) = (.Width - 1000) * 0.24
        .ColWidth(3) = (.Width - 1000) * 0.12
        .ColWidth(4) = (.Width - 1000) * 0.12
        .ColWidth(5) = (.Width - 1000) * 0.12
        .ColWidth(6) = (.Width - 1000) * 0.21
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        If .Cols > 10 Then .ColWidth(10) = 0
        If .Cols > 10 Then .ColWidth(11) = 0
        If .Cols > 10 Then .ColWidth(12) = 0
        pBar.Left = .Left
        pBar.Width = .Width
    End With
    
    lblHeading(4).Top = flxParts.Top - 255
    lblHeading(5).Top = flxParts.Top - 235
End Sub

Private Sub LoadSpareParts()
    Dim Qry As String
    Dim TtlRec As Integer
    Dim pPct As Single
        
        
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    
End Sub

Public Sub DisplaySPOrder(MyRef As String)
    Dim rsOrder As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, Plst As String
    Dim i As Integer, qOrd As Long, qRec As Long, TtlqRec As Long
    
    '****Access
    'Qry = "SELECT SPORDER.* From SPORDER " & _
          "WHERE (((SPORDER.REFERENCE)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT SPORDER.* From SPORDER " & _
          "WHERE SPORDER.REFERENCE='" & MyRef & "';"
    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Call ClearEntries
            .MoveFirst
            txtPODate.Text = !ORDERDATE
            txtPONumber.Text = !REFERENCE
            cboType.Text = !ORDERTYPE
            txtTerms.Text = IIf(IsNull(!Terms), "", !Terms)
            txtSuppID.Text = IIf(IsNull(!SUPPLIERID), "", !SUPPLIERID)
            txtOrderID.Text = !SPORDERID
        End If
        .Close
    End With
    If txtSuppID.Text <> "" Then Call DisplaySupplier
    
    For i = 3 To 5
        cmdTrans(i).Enabled = False
    Next i
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER, SPITR.PRICE, [ORDER]*[PRICE] AS TOTAL, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & MyRef & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER`, SPITR.PRICE, SPITR.`ORDER`*SPITR.PRICE AS TOTAL, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "' ORDER BY SPAREPARTS.PARTNUMBER;"
    
    qOrd = 0
    qRec = 0
    TtlqRec = 0
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxParts.Rows = 1
            .MoveFirst
            pBar.Visible = True
            pBar.Value = 0
            pBar.Max = .RecordCount
            Do While Not .EOF
                flxParts.AddItem flxParts.Rows & vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & vbTab & vbTab & vbTab & .Fields(5)
                qOrd = qOrd + !Order
                qRec = GetPOrec(MyRef, !SparepartID)
                TtlqRec = TtlqRec + qRec
                Plst = GetPackingList(MyRef, !SparepartID)
                flxParts.TextMatrix(flxParts.Rows - 1, 4) = qRec
                flxParts.TextMatrix(flxParts.Rows - 1, 5) = !Order - qRec
                flxParts.TextMatrix(flxParts.Rows - 1, 6) = Plst
                .MoveNext
                pBar.Value = pBar.Value + 1
                DoEvents
            Loop
        End If
        .Close
        pBar.Visible = False
    End With
    If qOrd = 0 And TtlqRec = 0 Then
        For i = 3 To 5
            cmdTrans(i).Enabled = True
        Next i
        Call LockData(True)
        Exit Sub
    End If
    txtOrder.Text = qOrd
    txtReceived.Text = TtlqRec
    lblRecPct.Caption = Format(TtlqRec / qOrd, "##0.00%") & " Served"
    txtBackOrder = qOrd - TtlqRec
    If qOrd - TtlqRec > 0 Then
        lblBckPct.Caption = Format((qOrd - TtlqRec) / qOrd, "##0.00%") & " Back Order"
    Else
        lblBckPct.Caption = "0.00% Back Order"
    End If
    'UPDATE PO
    '***Access
    'LYRcon.Execute "UPDATE SPORDER SET SPORDER.QTYORDER = " & qOrd & ", SPORDER.QTYREC = " & TtlqRec & " , SPORDER.BACKORDER = " & qOrd - TtlqRec & ", SPORDER.BOLASTUPDATE = #" & Date & "# " & _
                   "WHERE (((SPORDER.SPORDERID)=" & txtOrderID.Text & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPORDER SET SPORDER.QTYORDER = " & qOrd & ", SPORDER.QTYREC = " & TtlqRec & " , SPORDER.BACKORDER = " & qOrd - TtlqRec & ", SPORDER.BOLASTUPDATE = '" & Format(Date, "yyyy-mm-dd") & "' " & _
                   "WHERE SPORDER.SPORDERID=" & txtOrderID.Text & ";"
    
    For i = 3 To 5
        cmdTrans(i).Enabled = True
    Next i
    Call LockData(True)
                   
End Sub

Function GetPOrec(MyRef As String, MyID As Integer) As Integer
    Dim rsROrder As New ADODB.Recordset
    Dim Qry As String, Alloc As Integer, Recv As Integer
    
    'Qry = "SELECT Sum(SPITR.[RECEIVE]) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.PACKINGLIST)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyRef & "';"
    
    With rsROrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetPOrec = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            GetPOrec = 0
        End If
        .Close
    End With            'Order Allocation - [Previous Receive Parts + Current Receive Parts]
End Function

Function GetPackingList(MyRef As String, MyID As Integer) As String
    Dim rsPList As New ADODB.Recordset
    Dim Qry As String, MyPL As String
    
    'Qry = "SELECT SPITR.REFERENCE From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.PACKINGLIST)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT SPITR.REFERENCE From SPITR " & _
          "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyRef & "';"
    
    With rsPList
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        MyPL = ""
        If .RecordCount > 0 Then
            Do While Not .EOF
                MyPL = IIf(IsNull(!REFERENCE), "", IIf(MyPL <> "", MyPL & ", ", "") & !REFERENCE)
                .MoveNext
            Loop
        End If
        .Close
    End With
    GetPackingList = MyPL
End Function

Private Sub DisplaySupplier()
    Dim rsSupp As New ADODB.Recordset
    Dim Qry As String
    
    
    'Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER WHERE (((SUPPLIER.SUPPLIERID)=" & txtSuppID.Text & "));"
    'mySQL
    Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE, SUPPLIER.ADDRESS " & _
          "From SUPPLIER WHERE SUPPLIER.SUPPLIERID=" & txtSuppID.Text & ";"
    
    With rsSupp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Text = !SUPPLIERNAME
        End If
        .Close
    End With
    
End Sub

Private Sub txtPODate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPONumber.SetFocus
End Sub

Private Sub txtPONumber_GotFocus()
    If txtPONumber.Text <> "" Then
        txtPONumber.SelStart = 0
        txtPONumber.SelLength = Len(txtPONumber.Text)
    End If
End Sub

Private Sub txtPONumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboType.SetFocus
End Sub

Private Sub txtPONumber_LostFocus()
    Dim Qry As String
    
    'Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE (((SPORDER.REFERENCE)='" & txtPONumber.Text & "'));"
    
    'mySQL
    Qry = "SELECT SPORDER.REFERENCE From SPORDER WHERE SPORDER.REFERENCE='" & txtPONumber.Text & "';"
    
    If AddNew Then
        If CtrlNumExist(Qry) Then
            MsgBox "P.O. Number already exist!", vbExclamation + vbOKOnly, "Validation Error"
            Exit Sub
        End If
    End If
End Sub


Private Sub txtSaddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTerms.SetFocus
End Sub

Private Sub txtSCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSaddress.SetFocus
End Sub

Private Sub txtSearch_Change()
    cmdNext.Tag = 1
End Sub

Private Sub txtSearch_Click()
    cmdNext.Tag = 1
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdNext_Click
End Sub


Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPODate.SetFocus
End Sub

