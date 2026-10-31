VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SPBegBal 
   Caption         =   "Spare Parts Inventory Beginning Balance"
   ClientHeight    =   9105
   ClientLeft      =   825
   ClientTop       =   645
   ClientWidth     =   11805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtGR 
      Height          =   285
      Left            =   6840
      TabIndex        =   39
      Top             =   9840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   720
      TabIndex        =   38
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSrch 
      Height          =   1695
      Left            =   3720
      TabIndex        =   37
      Top             =   9720
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2990
      _Version        =   393216
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab tabFind 
      Height          =   1215
      Left            =   7080
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2143
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPBegBal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNext"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboMatch"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSearch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.ComboBox txtSearch 
         Height          =   315
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboMatch 
         Height          =   315
         ItemData        =   "SPBegBal.frx":001C
         Left            =   1080
         List            =   "SPBegBal.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   675
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3480
         TabIndex        =   31
         Top             =   645
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Find &Next"
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Match:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   705
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Find What:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   9240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   9240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   9240
   End
   Begin VB.TextBox txtFind 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Part Number Exist @ row 5"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   2640
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdAltA 
      Caption         =   "&A"
      Height          =   375
      Left            =   0
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   11040
      Width           =   375
   End
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Top             =   9000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cboPartNo 
      Height          =   315
      Left            =   1320
      TabIndex        =   16
      Top             =   9480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid flxBeg 
      Height          =   6495
      Left            =   360
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "XYF"
      Top             =   2400
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11456
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   $"SPBegBal.frx":0020
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5520
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":00AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":0649
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":0BE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":117D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":1717
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":1CB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":224B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":27E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":2D7F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   360
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   720
      Width           =   11295
      Begin VB.TextBox txtPreparedBy 
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
         Left            =   3120
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtInvDate 
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
         Left            =   9480
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtCtrlNo 
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtCheckedBy 
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
         Left            =   3120
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE ENCODED"
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
         Index           =   6
         Left            =   7440
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "INVENTORY CTRL. NO."
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
         Left            =   7320
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   630
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "INVENTORY REPORT CHECKED BY:"
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
         Left            =   120
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "INVENTORY REPORT PREPARED BY:"
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
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   270
         Width           =   3135
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":3319
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":41F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":50CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":5FA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SPBegBal.frx":6E81
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabProg 
      Height          =   735
      Left            =   3240
      TabIndex        =   18
      Top             =   9120
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1296
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPBegBal.frx":7D5B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLoading"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPct"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pBar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblPct 
         Alignment       =   1  'Right Justify
         Caption         =   "0%"
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblLoading 
         Caption         =   "Loading Previously Saved Spare Parts"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   3615
      End
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   285
      Left            =   9840
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   33
      Top             =   0
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&Find"
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
      Image           =   "SPBegBal.frx":7D77
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&Print"
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
      Image           =   "SPBegBal.frx":8311
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   35
      Top             =   0
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "SPBegBal.frx":88AB
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   36
      Top             =   0
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      Caption         =   "&Lock"
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
      Image           =   "SPBegBal.frx":8E45
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   6000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11925
      m_MDI_Height    =   9510
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "LIST OF SPAREPARTS"
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
      Height          =   300
      Index           =   4
      Left            =   360
      TabIndex        =   14
      Top             =   2025
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SPAREPARTS BEGINNING BALANCE"
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
      Height          =   300
      Index           =   2
      Left            =   7155
      TabIndex        =   12
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
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   10
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
      Height          =   300
      Index           =   1
      Left            =   390
      TabIndex        =   11
      Top             =   480
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SPAREPARTS BEGINNING BALANCE"
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
      Height          =   300
      Index           =   3
      Left            =   7185
      TabIndex        =   13
      Top             =   480
      Width           =   4500
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "LIST OF SPAREPARTS"
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
      Height          =   300
      Index           =   5
      Left            =   390
      TabIndex        =   15
      Top             =   2055
      Width           =   2415
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "FlxTrans"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row        Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
   Begin VB.Menu mnuDesc 
      Caption         =   "Description"
      Visible         =   0   'False
      Begin VB.Menu mnuEditDesc 
         Caption         =   "Edit Description"
      End
   End
End
Attribute VB_Name = "SPBegBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset

Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean, OrdSrch As Boolean
Dim ActRow As Integer, ActCol As Integer
Dim UpdatingRow As Integer

Function CanAddRow() As Boolean
    Dim i As Integer
    
    ActRow = flxBeg.Rows - 1
    If ActRow <> flxBeg.Rows - 1 Then
        CanAddRow = False
        Exit Function
    End If
    For i = 1 To 8          'loop for visible cols. dapat w/ data all cols xcept location
        If i <> 6 And i <> 4 Then
            If flxBeg.TextMatrix(ActRow, i) = "" Then
                CanAddRow = False
                Exit Function
            End If
        End If
    Next i
    If Not IsDate(flxBeg.TextMatrix(ActRow, 5)) Then
        CanAddRow = False
        Exit Function
    End If
    CanAddRow = True
End Function

Private Sub AddInCombo()
    Dim x As Integer, w As Boolean
    
    If txtSearch.Text = "" Then Exit Sub
    w = False
    For x = 0 To txtSearch.ListCount - 1
        If txtSearch.Text = txtSearch.List(x) Then
            If x = 0 Then
                w = True
                Exit For
            Else
                txtSearch.RemoveItem x
                Exit For
            End If
        End If
    Next x
    If w = False Then txtSearch.AddItem txtSearch.Text, 0
End Sub

Private Sub ClearEntries()
    txtPreparedBy.Text = ""
    txtCheckedBy.Text = ""
    txtInvDate.Text = ""
    txtCtrlNo.Text = ""
    flxBeg.Rows = 1
    flxBeg.Rows = 2
End Sub

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    flxBeg.Enabled = Not Bln
End Sub


Private Sub cboFind_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    With cboFind
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            If ActRow = i + 1 Then Exit Sub
            txtFind.Tag = i + 1 'IIf(LastRowLoaded >= i, i + 1, i)
            txtFind.Top = (flxBeg.CellTop + flxBeg.Top) - 315
            txtFind.Left = flxBeg.CellLeft + flxBeg.Left
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

Private Sub cboMatch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then cmdNext.SetFocus
End Sub

Private Sub cmdAltA_Click()
    Call mnuAdd_Click
End Sub

Private Sub LoadTemp()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer, TtlRec As Integer
    Dim Qry As String
    Dim pPct As Single
    
    With rsTemp

        Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.TRANSDATE, SPAREPARTS.LOCATION, SPITR.SOURCE, IIf(IsNull([SPAREPARTS].[REMARKS]),'',Left$([SPAREPARTS.REMARKS],1)) AS GR, SPAREPARTS.SPAREPARTID, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.STOREPRICE, SPAREPARTS.COST, SPAREPARTS.APPLICMODEL " & _
              "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.TRANSTYPE)='BEGBAL')) ORDER BY SPAREPARTS.PARTNUMBER;"

        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxBeg.Rows = 1
            'Start Progress bar
            tabProg.Top = 3000
            tabProg.Caption = ""
            tabProg.Visible = True
            tabProg.ZOrder
            For i = 0 To 3
                cmdTrans(i).Enabled = False
            Next i
            pBar.Value = 0
            pBar.Max = .RecordCount
            lblLoading.Caption = "Loading Spare Parts"
            Do While Not .EOF
                flxBeg.AddItem vbTab & .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2) & vbTab & .Fields(3) & vbTab & Format(.Fields(4), "mm/dd/yy") & _
                vbTab & .Fields(5) & vbTab & IIf(IsNull(.Fields(6)), "", .Fields(6)) & vbTab & IIf(IsNull(.Fields(7)), "", .Fields(7)) & vbTab & .Fields(8) & vbTab & .Fields(9)
                cboFind.AddItem .Fields(1)
                pBar.Value = pBar.Value + 1
                pPct = (pBar.Value / pBar.Max)
                DoEvents
                lblPct.Caption = Format$(pPct, "##0.0%")
                .MoveNext
            Loop
            For i = 0 To 3
                cmdTrans(i).Enabled = True
            Next i
            tabProg.Visible = False
            DoEvents
            tabProg.Top = 12560
            flxBeg.Enabled = True
            flxBeg.SetFocus
            flxBeg.Row = flxBeg.Rows - 1
            ActRow = flxBeg.Rows - 1
            flxBeg.Col = 0
'            LastRowLoaded = ActRow
            If ActRow > 18 Then flxBeg.Width = 11295
        End If
        .Close
    End With
    
End Sub

Private Sub SaveTemp()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    
    On Error GoTo LocalErr
    txtFlx.Visible = False
    txtPartNo.Visible = False
    txtDate.Visible = False
    ActRow = flxBeg.Rows - 1
    With rsTemp
        .Open "Temp", LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                .Delete 'adAffectAllChapters
                .MoveNext
            Loop
            '.Delete adAffectGroup
        End If
        If Not CanAddRow Then flxBeg.RemoveItem flxBeg.Rows - 1
        For i = 1 To flxBeg.Rows - 1
            .AddNew
            For j = 1 To 9
                .Fields(j - 1) = flxBeg.TextMatrix(i, j)
            Next j
            .Update
        Next i
        .Close
    End With
    MsgBox "Spare Parts temporarily saved!", vbInformation + vbOKOnly, "Save Temp."
    Exit Sub
LocalErr:
    If Err.Number = 381 Then Resume Next
        
End Sub

Private Sub cmdCancel_Click()
    tabFind.Visible = False
    cmdNext.Tag = 1
    cmdCancel.Tag = "hide"
    Call AddInCombo
    flxBeg.SetFocus
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call cmdCancel_Click
End Sub

Private Sub cmdCancel_LostFocus()
    If cmdCancel.Tag = "" Then txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    Dim i As Integer, j As Integer
    
    With flxBeg
        For i = Val(cmdNext.Tag) To .Rows - 1
            For j = 1 To 7
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
                        .SetFocus: .Row = i: .Col = 0: .ColSel = 8
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
    Dim Qry As String, myMsg As String
    Dim Ans As Integer
    
    Select Case Index
        Case 0  '"print"
            PRNmc.MySource = "SPBEGINBAL"
            PRNmc.Show
        Case 1  '"find"
            tabFind.Top = flxBeg.Top - tabFind.Height + 50
            tabFind.Left = (flxBeg.Left + flxBeg.Width) - tabFind.Width
            tabFind.Visible = True
            cboMatch.Text = cboMatch.List(1)
            cmdNext.Tag = 1
            cmdCancel.Tag = ""
            If Len(txtSearch.Text) > 0 Then
                txtSearch.SelStart = 0
                txtSearch.SelLength = Len(txtSearch.Text)
            End If
            txtSearch.SetFocus
        Case 2  '"final"
            myMsg = ValidateBBal
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            txtPartNo.Visible = False
            txtFlx.Visible = False
            txtDate.Visible = False
            Ans = MsgBox("This will LOCK the editing of Spare Parts beginning balance" & vbCrLf & _
                "Do you want to continue?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            txtCtrlNo.SetFocus
            Call SaveBegDesc(3, "True")
            Frame1.Enabled = False
            flxBeg.Tag = "just scroll"
            MsgBox "Beginning Balance successfully Locked!", vbInformation + vbOKOnly, "Final"
        Case 3  '"quit"
            MotoRep.tbMain.Visible = True
            Unload Me
    End Select

End Sub

Private Sub flxbeg_Click()
    Dim i As Integer, Tmp As String
    ActRow = flxBeg.Row
    ActCol = flxBeg.Col
    
    If ActCol = 0 Or ActRow = 0 Or flxBeg.Tag = "just scroll" Then Exit Sub
    Select Case ActCol
        Case 2  'PartNumber
            If flxBeg.TextMatrix(ActRow, 2) <> "" Then
                For i = 2 To 13
                    Tmp = IIf(Tmp <> "", Tmp & "|" & flxBeg.TextMatrix(ActRow, i), flxBeg.TextMatrix(ActRow, i))
                Next i
                flxBeg.TextMatrix(ActRow, 14) = Tmp
                txtPartNo.Tag = "exist"
            Else
                txtPartNo.Tag = ""
            End If
            With txtPartNo
                .Top = flxBeg.CellTop + flxBeg.Top
                .Left = flxBeg.CellLeft + flxBeg.Left
                .Width = flxBeg.CellWidth
                .Text = flxBeg.TextMatrix(ActRow, ActCol)
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 3  'Description can be edited only if part no. is not available in price list
            If flxBeg.TextMatrix(ActRow, 11) <> "eDesc" Then Exit Sub
            If flxBeg.TextMatrix(ActRow, ActCol) <> "" Then txtFlx.Tag = flxBeg.TextMatrix(ActRow, ActCol)
            With txtFlx
                .Top = flxBeg.CellTop + flxBeg.Top
                .Left = flxBeg.CellLeft + flxBeg.Left
                .Width = flxBeg.CellWidth
                .Text = flxBeg.TextMatrix(ActRow, ActCol)
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 8  'G/R
            If flxBeg.TextMatrix(ActRow, ActCol) <> "" Then txtGR.Tag = flxBeg.TextMatrix(ActRow, ActCol)
            With txtGR
                .Top = flxBeg.CellTop + flxBeg.Top
                .Left = flxBeg.CellLeft + flxBeg.Left
                .Width = flxBeg.CellWidth
                .Text = flxBeg.TextMatrix(ActRow, ActCol)
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        
        Case 1, 4, 6, 7 'Qty. and Remarks
            If flxBeg.TextMatrix(ActRow, ActCol) <> "" Then txtFlx.Tag = flxBeg.TextMatrix(ActRow, ActCol)
            With txtFlx
                .Top = flxBeg.CellTop + flxBeg.Top
                .Left = flxBeg.CellLeft + flxBeg.Left
                .Width = flxBeg.CellWidth
                .Text = flxBeg.TextMatrix(ActRow, ActCol)
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 5  'Date
            txtDate.Tag = IIf(flxBeg.TextMatrix(ActRow, ActCol) <> "", flxBeg.TextMatrix(ActRow, ActCol), "")
            With txtDate
                .Top = flxBeg.CellTop + flxBeg.Top
                .Left = flxBeg.CellLeft + flxBeg.Left
                .Width = flxBeg.CellWidth
                .Text = IIf(flxBeg.TextMatrix(ActRow, ActCol) <> "", flxBeg.TextMatrix(ActRow, ActCol), "__/__/__")
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        End Select
End Sub


Private Sub flxbeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxbeg_Click
End Sub

Private Sub flxbeg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case ActCol
        Case 0
            If Button = vbRightButton Then
                Call PopupMenu(mnuTrans)
            End If
        Case 3
            If Button = vbRightButton Then
                Call PopupMenu(mnuDesc)
            End If
    End Select
End Sub

Private Sub flxBeg_Scroll()
    txtPartNo.Visible = False
    txtFlx.Visible = False
    txtDate.Visible = False
    txtGR.Visible = False
End Sub

Private Sub flxSrch_Click()
    FromFlx = True
End Sub

Private Sub flxSrch_EnterCell()
    With flxSrch
        FromFlx = True
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        txtPartNo.Text = .TextMatrix(.Row, 1)
    End With
End Sub

Private Sub flxSrch_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Orig() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        If txtFind.Tag <> "" Then
            txtFind.Visible = False
            For i = 2 To 15
                flxBeg.TextMatrix(ActRow, i) = ""
            Next i
        End If
        If flxBeg.TextMatrix(ActRow, 14) <> "" Then
            Orig = Split(flxBeg.TextMatrix(ActRow, 14), "|")
            For i = 2 To 13
                flxBeg.TextMatrix(ActRow, i) = Orig(i - 2)
            Next i
            flxBeg.TextMatrix(ActRow, 14) = ""
        Else
            For i = 2 To 15
                flxBeg.TextMatrix(ActRow, i) = ""
            Next i
        End If
        txtPartNo.Tag = "SkipLostFocus"
        txtPartNo.Visible = False
        flxBeg.SetFocus
        flxBeg.Col = 1
    ElseIf KeyCode = vbKeyDown Then
        
    End If
End Sub

Private Sub flxSrch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        flxBeg.SetFocus
        flxBeg.Col = 4
    End If
End Sub

Private Sub flxSrch_LostFocus()
    flxSrch.Visible = False
End Sub


Private Sub Form_Load()
    Dim rsBBal As New ADODB.Recordset
    Dim rsBegDesc As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    
    'For i = TBcmd.LBound To TBcmd.UBound
    '    TBcmd(i).Top = MOTOREP.ScaleHeight + 1000
    'Next i
    cboMatch.AddItem "Whole Word"
    cboMatch.AddItem "Any Part of Word"
    tabFind.Caption = ""
    UpdatingRow = 0
    Qry = "SELECT BEGINDESC.* From BEGINDESC " & _
          "WHERE (((BEGINDESC.BEGCODE)='SP'));"
    With rsBegDesc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtPreparedBy.Text = IIf(IsNull(.Fields("PreparedBy")), "", .Fields("PreparedBy"))
            txtCheckedBy.Text = IIf(IsNull(.Fields("CheckedBy")), "", .Fields("CheckedBy"))
            txtInvDate.Text = IIf(IsNull(.Fields("InvDate")), Date, .Fields("InvDate"))
            txtCtrlNo.Text = .Fields("CtrlNo")
            If .Fields("Locked") = True Then
                Timer3.Enabled = True
                Exit Sub
            End If
        End If
        .Close
    End With
            
    Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PACKINGLIST, SPITR.TRANSDATE, SPAREPARTS.LOCATION, SPITR.SOURCE, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.STOREPRICE, SPAREPARTS.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID Where (((SPITR.TransType) = 'BEGBAL')) ORDER BY SPAREPARTS.PARTNUMBER;"

    With rsBBal
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            Timer1.Enabled = True
            Call LockData(False)
        Else
            Timer2.Enabled = True
            Call LockData(False)
        End If
        .Close
    End With
    MotoRep.sBar.Panels(6).Text = "Spare Parts Inventory Beginning Balance"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub


Private Sub mnuAdd_Click()
    If Not CanAddRow Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        Exit Sub
    Else
        If flxBeg.TextMatrix(flxBeg.Rows - 1, 10) = "" Then
            Call AddRecord
            cboFind.AddItem flxBeg.TextMatrix(ActRow, 2)
            flxBeg.SetFocus
            flxBeg.Col = 1
        Else
            flxBeg.AddItem ""
            If flxBeg.Rows > 18 Then flxBeg.Width = 11295
            flxBeg.SetFocus
            flxBeg.Col = 1
        End If
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    Dim myID As Long
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        myID = IIf(flxBeg.TextMatrix(ActRow, 9) <> "", Val(flxBeg.TextMatrix(ActRow, 9)), 0)
        txtFlx.Visible = False
        txtPartNo.Visible = False
        If flxBeg.Rows = 2 Then
            flxBeg.Rows = 1
            flxBeg.Rows = 2
        Else
            flxBeg.RemoveItem ActRow
        End If
        ActRow = flxBeg.Rows - 1
        flxBeg.Row = flxBeg.Rows - 1
        If myID <> 0 Then Call DeleteSP(myID)
        UpdatingRow = 0
    End If
End Sub

Private Sub mnuEditDesc_Click()
    If flxBeg.TextMatrix(ActRow, ActCol) <> "" Then txtFlx.Tag = flxBeg.TextMatrix(ActRow, ActCol)
    With txtFlx
        .Top = flxBeg.CellTop + flxBeg.Top
        .Left = flxBeg.CellLeft + flxBeg.Left
        .Width = flxBeg.CellWidth
        .Text = flxBeg.TextMatrix(ActRow, ActCol)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With

End Sub


Private Sub LoadSpareParts()
    Dim Qry As String, i As Integer
        
    
    Qry = "SELECT SOURCESP.PARTNUMBER, SOURCESP.DESCRIPTION, SOURCESP.SUPERSEDING, SOURCESP.COST, SOURCESP.RETAILPRICE, SOURCESP.APPLICMODEL " & _
          "From SOURCESP ORDER BY SOURCESP.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSrch.DataSource = rsParts
    flxSrch.ColWidth(0) = 0
    flxSrch.ColWidth(1) = txtPartNo.Width
    For i = 2 To flxSrch.Cols - 1
        flxSrch.ColWidth(i) = 0
    Next i

End Sub


Function ValidateBBal() As String
    Dim vMsg As String
    Dim i As Integer, cnt As Integer
    
    vMsg = ""
    If txtPreparedBy.Text = "" Then vMsg = "The field Prepared By must not be empty."
    If txtCtrlNo.Text = "" Then vMsg = vMsg & vbCrLf & "Field Inventory Ctrl. No. is required."
    If Not IsDate(txtInvDate.Text) Then vMsg = vMsg & vbCrLf & "Please enter a valid date."
    If txtCheckedBy.Text = "" Then vMsg = vMsg & vbCrLf & "Field Checked By must not be empty."
    cnt = 0
    For i = 1 To 3
        If flxBeg.TextMatrix(flxBeg.Rows - 1, i) = "" Then Cntr = Cntr + 1
    Next i
    If cnt > 0 Then vMsg = vMsg & vbCrLf & "Incomplete entry on last row of list of Units."
    ValidateBBal = vMsg
    
End Function

Private Sub SaveBBal()
    Dim rsTrans As New ADODB.Recordset
    Dim rsMotor As New ADODB.Recordset
    Dim rsBeginDesc As New ADODB.Recordset
    
    Dim MyErrlvl As String
    Dim i As Integer, myID As Long
    
'    On Error GoTo LocalErr
        
    MyErrlvl = "Saving MC Beginning Balance Description"
    With rsBeginDesc
        .Open "BeginDesc", LYRcon, adOpenKeyset, adLockOptimistic
        .AddNew
        .Fields("BegCode") = "SP"
        .Fields("PreparedBy") = txtPreparedBy.Text
        .Fields("CheckedBy") = txtCheckedBy.Text
        .Fields("InvDate") = txtInvDate.Text
        .Fields("CtrlNo") = txtCtrlNo.Text
        .Fields("TransNo") = "BBAL-" & txtCtrlNo.Text
        .Fields("UpdatedBy") = UserID
        .Fields("LastUpdate") = Date
        .Update
        .Close
    End With
    
    MyErrlvl = "Saving table SpareParts"
    With rsTrans
        .Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
        For i = 1 To flxBeg.Rows - 1
            .AddNew
            .Fields("PartNumber") = flxBeg.TextMatrix(i, 2)
            .Fields("Description") = flxBeg.TextMatrix(i, 3)
            .Fields("SupersedingPN") = flxBeg.TextMatrix(i, 8)
            .Fields("StorePrice") = IIf(IsNumeric(flxBeg.TextMatrix(i, 9)), flxBeg.TextMatrix(i, 9), 0)
            .Fields("QtyOnHand") = Val(flxBeg.TextMatrix(i, 1))
            .Fields("BegBal") = Val(flxBeg.TextMatrix(i, 1))
            .Fields("Location") = GetLocation(flxBeg.TextMatrix(i, 6))
            .Fields("DateBegBal") = Date
            .Fields("UpdatedBy") = UserID
            .Fields("LastUpdate") = Date
            .Update
            flxBeg.TextMatrix(i, 10) = .Fields("SparePartID")
        Next i
        .Close
    End With
    
    MyErrlvl = "Saving table MC Transaction Register"
    With rsTrans
        .Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
        For i = 1 To flxBeg.Rows - 1
            .AddNew
            .Fields("TransNo") = "BBAL-" & txtCtrlNo.Text
            .Fields("TransDate") = flxBeg.TextMatrix(i, 5)
            .Fields("SparePartID") = flxBeg.TextMatrix(i, 10)
            .Fields("Received") = Val(flxBeg.TextMatrix(i, 1))
            .Fields("PackingList") = flxBeg.TextMatrix(i, 4)
            .Fields("Source") = flxBeg.TextMatrix(i, 7)
            .Update
        Next i
        .Close
    End With
        
'    Exit Sub
'LocalErr:
'    MsgBox MyErrlvl & vbCrLf & str(Err.Number) & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Function GetLocation(MyPart As String) As String
    Dim rsLoc As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT Sheet1.PARTNUMBER, Sheet1.LOCATION From Sheet1 " & _
          "WHERE (((Sheet1.PARTNUMBER)='" & MyPart & "'));"
    With rsLoc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetLocation = IIf(IsNull(.Fields("Location")), "", .Fields("Location"))
        Else
            GetLocation = ""
        End If
        .Close
    End With
End Function


Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    With flxBeg
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 600) * 0.08
        .ColWidth(2) = (.Width - 600) * 0.17
        .ColWidth(3) = (.Width - 600) * 0.2
        .ColWidth(4) = (.Width - 600) * 0.1
        .ColWidth(5) = (.Width - 600) * 0.1
        .ColWidth(6) = (.Width - 600) * 0.15
        .ColWidth(7) = (.Width - 600) * 0.15
        .ColWidth(8) = (.Width - 600) * 0.05
        For i = 9 To 15
            .ColWidth(i) = 0
        Next i
        lblHeading(4).Left = .Left
        lblHeading(5).Left = .Left + 30
        lblHeading(4).Top = .Top - lblHeading(4).Height + 30
        lblHeading(5).Top = .Top - lblHeading(5).Height
    End With
    lblHeading(0).Left = Frame1.Left
    lblHeading(1).Left = Frame1.Left + 30
    lblHeading(0).Top = Frame1.Top - lblHeading(0).Height + 30
    lblHeading(1).Top = Frame1.Top - lblHeading(1).Height
    lblHeading(2).Left = Frame1.Left + Frame1.Width - lblHeading(2).Width
    lblHeading(3).Left = Frame1.Left + Frame1.Width - lblHeading(3).Width
    lblHeading(2).Top = lblHeading(0).Top
    lblHeading(3).Top = lblHeading(1).Top
    
    
End Sub

Private Sub Timer1_Timer()
    ' Call tbSaveCancel(True)
    Call LoadTemp
    Call LoadSpareParts
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Call LoadSpareParts
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    Call LoadTemp
    Frame1.Enabled = False
    flxBeg.Tag = "just scroll"
    Timer3.Enabled = False
End Sub

Private Sub txtCheckedBy_GotFocus()
    txtCheckedBy.Tag = txtCheckedBy.Text
End Sub

Private Sub txtCheckedBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCheckedBy_LostFocus()
    With txtCheckedBy
        If .Text <> "" And .Tag <> .Text Then Call SaveBegDesc(1, .Text)
    End With
End Sub

Private Sub txtCtrlNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxBeg.SetFocus
        flxBeg.Col = 0
    End If
End Sub


Private Sub txtFind_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call FixDoublePartNo
End Sub

Private Sub FixDoublePartNo()
    Dim i As Integer
        
    txtFind.Visible = False
    txtPartNo.Tag = "SkipLostFocus"
    txtPartNo.Visible = False
    For i = 0 To 15
        flxBeg.TextMatrix(ActRow, i) = ""
    Next i
    flxBeg.Row = Val(txtFind.Tag)
    flxBeg.Col = 0: flxBeg.ColSel = 9
    flxBeg.TopRow = Val(txtFind.Tag)
    flxBeg.SetFocus
End Sub


Private Sub txtFlx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       flxBeg.TextMatrix(ActRow, ActCol) = txtFlx.Tag
       txtFlx.Tag = "skip"
       txtFlx.Text = ""
       txtFlx.Visible = False
       flxBeg.SetFocus
       flxBeg.Col = ActCol
    End If
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If ActCol = 1 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    ElseIf ActCol = 3 Or ActCol = 4 Or ActCol = 6 Or ActCol = 7 Then
        KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    End If
    
    If KeyAscii = 13 Then
        Call txtFlx_LostFocus
        flxBeg.SetFocus
        flxBeg.Col = ActCol + 1
    End If
End Sub

Private Sub txtFlx_LostFocus()
    Dim Ans As Integer
    
    If txtFlx.Tag = "skip" Then
        txtFlx.Tag = ""
        Exit Sub
    End If
    If txtFlx.Tag <> txtFlx.Text Then   'Or ActCol = 6 Then
        If txtFlx.Text <> txtFlx.Tag And flxBeg.TextMatrix(ActRow, 9) <> "" And UpdatingRow <> ActRow Then
            Ans = MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, "Save")
            If Ans = vbYes Then
                flxBeg.TextMatrix(ActRow, ActCol) = txtFlx.Text
                Select Case ActCol
                    Case 1  'Quantity
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = " & flxBeg.TextMatrix(ActRow, ActCol) & ", SPAREPARTS.BEGBAL = " & flxBeg.TextMatrix(ActRow, ActCol) & _
                                       " WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
                        LYRcon.Execute "UPDATE SPITR SET SPITR.RECEIVE = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPITR.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"
                    Case 3  'Description
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
                    Case 4  'Packing List->SPITR [SpareParts->Reference=BEG. BAL.
                        LYRcon.Execute "UPDATE SPITR SET SPITR.PACKINGLIST = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPITR.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"
                    Case 6  'Location
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
                    Case 7  'REC FROM
                        LYRcon.Execute "UPDATE SPITR SET SPITR.SOURCE = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPITR.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"
                End Select
                txtFlx.Tag = ""
                txtFlx.Visible = False
                Exit Sub
            Else
                flxBeg.TextMatrix(ActRow, ActCol) = txtFlx.Tag
                txtFlx.Tag = ""
                txtFlx.Visible = False
                Exit Sub
            End If
        ElseIf txtFlx.Text <> txtFlx.Tag And flxBeg.TextMatrix(ActRow, 9) <> "" And UpdatingRow = ActRow Then
            flxBeg.TextMatrix(ActRow, ActCol) = txtFlx.Text
                Select Case ActCol
                    Case 1  'Quantity
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = " & flxBeg.TextMatrix(ActRow, ActCol) & ", SPAREPARTS.BEGBAL = " & flxBeg.TextMatrix(ActRow, ActCol) & _
                                       " WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
                    Case 3  'Description
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
                    Case 4  'Packing List->SPITR [SpareParts->Reference=BEG. BAL.
                        LYRcon.Execute "UPDATE SPITR SET SPITR.PACKINGLIST = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPITR.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"
                    Case 6  'Location
                        LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
                    Case 7  'REC FROM
                        LYRcon.Execute "UPDATE SPITR SET SPITR.SOURCE = '" & flxBeg.TextMatrix(ActRow, ActCol) & "' " & _
                                       "WHERE (((SPITR.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"
                End Select
                txtFlx.Tag = ""
                txtFlx.Visible = False
                Exit Sub
        End If
    End If
    txtFlx.Tag = ""
    flxBeg.TextMatrix(ActRow, ActCol) = txtFlx.Text
    txtFlx.Visible = False
    If CanAddRow Then
        If flxBeg.TextMatrix(ActRow, 9) = "" Then
            Call AddRecord
            cboFind.AddItem flxBeg.TextMatrix(ActRow, 2)
        Else
            If ActRow = flxBeg.Rows - 1 Then
                flxBeg.AddItem ""
            End If
        End If
    End If
End Sub

Private Sub txtGR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       flxBeg.TextMatrix(ActRow, 8) = txtGR.Tag
       txtGR.Tag = "skiplostfocus"
       txtGR.Text = ""
       txtGR.Visible = False
       flxBeg.SetFocus
       flxBeg.Col = ActCol
    End If
End Sub

Private Sub txtGR_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "GR"
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtGR_LostFocus
        flxBeg.SetFocus
        flxBeg.Col = 1
        If ActRow < flxBeg.Rows - 1 Then
            flxBeg.Row = ActRow + 1
        End If
    End If
End Sub

Private Sub txtGR_LostFocus()
    Dim Ans As Integer
    
    If txtGR.Tag = "skiplostfocus" Then
        txtGR.Tag = ""
        Exit Sub
    End If
    'Update Remarks to Genuine
    If txtGR.Text <> txtGR.Tag And flxBeg.TextMatrix(ActRow, 9) <> "" And UpdatingRow = ActRow Then
        flxBeg.TextMatrix(ActRow, ActCol) = txtGR.Text
        If txtGR.Text = "G" Then
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.REMARKS = 'GENUINE' WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
        ElseIf txtGR.Text = "R" Then
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.REMARKS = 'REPLACEMENT' WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
        End If
        txtGR.Tag = ""
        txtGR.Visible = False
        Exit Sub
    End If
    txtGR.Tag = ""
    flxBeg.TextMatrix(ActRow, ActCol) = txtGR.Text
    txtGR.Visible = False
    If CanAddRow Then
        If flxBeg.TextMatrix(ActRow, 9) = "" Then
            Call AddRecord
            cboFind.AddItem flxBeg.TextMatrix(ActRow, 2)
        Else
            If ActRow = flxBeg.Rows - 1 Then
                flxBeg.AddItem ""
            End If
        End If
    End If

End Sub

Private Sub txtInvDate_GotFocus()
    txtInvDate.Tag = txtInvDate.Text
End Sub

Private Sub txtInvDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtInvDate_LostFocus()
    With txtInvDate
        If IsDate(.Text) And .Tag <> .Text Then Call SaveBegDesc(2, .Text)
    End With
End Sub

Private Sub txtPartNo_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    With flxSrch
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            If Not FromFlx Then     'top row only if user type @ txtpartno
                .TopRow = rsParts.AbsolutePosition - 1
                FromFlx = False
            End If
            .Row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 6
            txtPartNo.Text = rsParts!PARTNUMBER
            'Insert Additional retrieval here
            flxBeg.TextMatrix(ActRow, 3) = .TextMatrix(.Row, 2) 'Description
            flxBeg.TextMatrix(ActRow, 7) = "BINOS MKTG."
            flxBeg.TextMatrix(ActRow, 8) = "G" 'Genuine
            flxBeg.TextMatrix(ActRow, 10) = .TextMatrix(.Row, 3) 'Superseding
            flxBeg.TextMatrix(ActRow, 11) = .TextMatrix(.Row, 5) 'Retail Price
            flxBeg.TextMatrix(ActRow, 12) = .TextMatrix(.Row, 4) 'Cost
            flxBeg.TextMatrix(ActRow, 13) = .TextMatrix(.Row, 6) 'Applicable Model
            If ActRow > 1 Then
                flxBeg.TextMatrix(ActRow, 4) = flxBeg.TextMatrix(ActRow - 1, 4)
                flxBeg.TextMatrix(ActRow, 5) = flxBeg.TextMatrix(ActRow - 1, 5)
                flxBeg.TextMatrix(ActRow, 7) = flxBeg.TextMatrix(ActRow - 1, 7)
            End If
                        
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            flxBeg.TextMatrix(ActRow, 2) = ""     'Description
            For i = 3 To 15      'Clear old values of FLXBEG
                flxBeg.TextMatrix(ActRow, i) = ""
            Next i
            .Visible = False
        End If
    End With
    
    cboFind.Text = txtPartNo.Text

End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Orig() As String
    Dim i As Integer
    
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        For i = 3 To 15
            flxBeg.TextMatrix(ActRow, i) = ""
        Next i
        cboFind.Text = ""
        cboFind.Text = txtPartNo.Text
        PrtSrch = True
    ElseIf KeyCode = vbKeyEscape Then
        If txtFind.Tag <> "" Then
            txtFind.Visible = False
            For i = 2 To 15
                flxBeg.TextMatrix(ActRow, i) = ""
            Next i
        End If
        If flxBeg.TextMatrix(ActRow, 14) <> "" Then
            Orig = Split(flxBeg.TextMatrix(ActRow, 14), "|")
            For i = 2 To 13
                flxBeg.TextMatrix(ActRow, i) = Orig(i - 2)
            Next i
            flxBeg.TextMatrix(ActRow, 14) = ""
        Else
            For i = 2 To 15
                flxBeg.TextMatrix(ActRow, i) = ""
            Next i
        End If
        txtPartNo.Tag = "SkipLostFocus"
        txtPartNo.Visible = False
        flxBeg.SetFocus
        flxBeg.Col = 1
    ElseIf KeyCode = vbKeyDown Then
        txtPartNo.Tag = "SkipLostFocus"
        flxSrch.Visible = True
        flxSrch.SetFocus
    Else
        FromFlx = False
    End If

End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtFind.Tag <> "" Then
            Call FixDoublePartNo
            Exit Sub
        End If
        Call txtPartNo_LostFocus
        flxBeg.SetFocus
        If flxBeg.TextMatrix(ActRow, 15) = "eDesc" Then
            flxBeg.Col = 3
        ElseIf txtPartNo.Tag = "No" Then
            flxBeg.Col = 1
        Else
            flxBeg.Col = 4
        End If
    End If

End Sub

Private Sub txtPartNo_LostFocus()
    Dim Ans As Integer
    Dim i As Integer
    Dim Orig() As String
    
    If txtFind.Tag <> "" Then
        Call FixDoublePartNo
        Exit Sub
    End If

    If txtPartNo.Tag = "SkipLostFocus" Then
        txtPartNo.Tag = ""
        Exit Sub
    End If
    If flxBeg.TextMatrix(ActRow, 3) = "" And Trim(txtPartNo.Text) <> "" Then
        Ans = MsgBox("Part Number does not exist in Price List!" & vbCrLf & "Do you want to proceed?", vbQuestion + vbYesNo, "Warning")
        If Ans = vbNo Then
            For i = 2 To 13
                flxBeg.TextMatrix(ActRow, i) = ""
            Next i
            If flxBeg.TextMatrix(ActRow, 14) <> "" Then
                Orig = Split(flxBeg.TextMatrix(ActRow, 14), "|")
                For i = 2 To 13
                    flxBeg.TextMatrix(ActRow, i) = Orig(i - 2)
                Next i
                flxBeg.TextMatrix(ActRow, 14) = ""
            End If
            txtPartNo.Tag = "No"
            flxBeg.TextMatrix(ActRow, 2) = ""
            txtPartNo.Visible = False
            flxSrch.Visible = False
            Exit Sub
        Else
            flxBeg.TextMatrix(ActRow, 15) = "eDesc"
            flxBeg.TextMatrix(ActRow, 2) = UCase$(txtPartNo.Text)
            flxBeg.TextMatrix(ActRow, 7) = "BINOS MKTG."
            If ActRow > 1 Then
                flxBeg.TextMatrix(ActRow, 4) = flxBeg.TextMatrix(ActRow - 1, 4)
                flxBeg.TextMatrix(ActRow, 5) = flxBeg.TextMatrix(ActRow - 1, 5)
                flxBeg.TextMatrix(ActRow, 7) = flxBeg.TextMatrix(ActRow - 1, 7)
            End If
            txtPartNo.Visible = False
            flxSrch.Visible = False
            Exit Sub
        End If
    End If
    If flxBeg.TextMatrix(ActRow, 14) <> "" And flxBeg.TextMatrix(ActRow, 2) <> txtPartNo.Text And flxBeg.TextMatrix(ActRow, 9) <> "" Then
        Ans = MsgBox("Do you want to apply changes made to Part No.?", vbQuestion + vbYesNo, "Apply Changes")
        If Ans = vbYes Then
            flxBeg.TextMatrix(ActRow, 2) = txtPartNo.Text
            'Update PartNumber and Description
            Call UpdatePartNo(flxBeg.TextMatrix(ActRow, 9))
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.PARTNUMBER = '" & txtPartNo.Text & "', SPAREPARTS.DESCRIPTION = '" & flxBeg.TextMatrix(ActRow, 3) & "' " & _
                           "WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"   'SpareParts ID
            txtPartNo.Visible = False
            flxSrch.Visible = False
        Else
            txtPartNo.Visible = False
            flxSrch.Visible = False
        End If
        Exit Sub
    End If
    flxBeg.TextMatrix(ActRow, 2) = txtPartNo.Text
    txtPartNo.Visible = False
    flxSrch.Visible = False
    UpdatingRow = ActRow
    If CanAddRow And flxBeg.TextMatrix(flxBeg.Rows - 1, 9) = "" Then
        cboFind.AddItem flxBeg.TextMatrix(ActRow, 2)
        Call AddRecord
    End If

End Sub

Private Sub txtPreparedBy_GotFocus()
    txtPreparedBy.Tag = txtPreparedBy.Text
End Sub

Private Sub txtPreparedBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxBeg.TextMatrix(ActRow, ActCol) = txtDate.Tag
        txtDate.Text = "__/__/__"
        txtDate.Visible = False
        txtDate.Tag = "SkipLostFocus"
        flxBeg.SetFocus
        flxBeg.Col = ActCol
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "Invalid date entered!", vbExclamation + vbOKOnly, "Validation Error"
            txtDate.SelStart = 1
            txtDate.SelLength = Len(txtDate.Text)
            Exit Sub
        End If
        If DateValue(txtDate.Text) < #1/1/2005# Or DateValue(txtDate.Text) > Date Then
            MsgBox "Invalid date entered!", vbExclamation + vbOKOnly, "Validation Error"
            txtDate.SelStart = 1
            txtDate.SelLength = Len(txtDate.Text)
            txtDate.SetFocus
            Exit Sub
        End If
        Call txtDate_LostFocus
        flxBeg.SetFocus
        flxBeg.Col = 6
    End If
End Sub

Private Sub txtDate_LostFocus()
    
    If txtDate.Tag = "SkipLostFocus" Then
        txtDate.Tag = ""
        Exit Sub
    End If
    If Not IsDate(txtDate.Text) Then
        MsgBox "Invalid date entered!", vbExclamation + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    If DateValue(txtDate.Text) < #1/1/2005# Or DateValue(txtDate.Text) > Date Then
        MsgBox "Invalid date entered!", vbExclamation + vbOKOnly, "Validation Error"
        txtDate.SelStart = 1
        txtDate.SelLength = Len(txtDate.Text)
        Exit Sub
    End If
    If txtDate.Tag <> "" Then
        If txtDate.Tag <> txtDate.Text And flxBeg.TextMatrix(ActRow, 9) <> "" And UpdatingRow <> ActRow Then
            Ans = MsgBox("Do you want to apply changes?", vbQuestion + vbYesNo, "Apply Changes")
            If Ans = vbYes Then
                flxBeg.TextMatrix(ActRow, ActCol) = txtDate.Text
                LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSDATE = #" & txtDate.Text & "# " & _
                               "WHERE (((SPITR.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DATEBEGBAL = #" & txtDate.Text & "# " & _
                               "WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
                'Call Edit457(flxBeg.TextMatrix(ActRow, 10), txtDate.Text, ActCol)
                txtDate.Visible = False
                txtDate.Tag = ""
                Exit Sub
            Else
                flxBeg.TextMatrix(ActRow, ActCol) = txtDate.Tag
                txtDate.Visible = False
                txtDate.Tag = "SkipLostFocus"
                Exit Sub
            End If
        ElseIf txtDate.Tag <> txtDate.Text And flxBeg.TextMatrix(ActRow, 9) <> "" And UpdatingRow = ActRow Then
            flxBeg.TextMatrix(ActRow, ActCol) = txtDate.Text
            LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSDATE = #" & txtDate.Text & "# " & _
                           "WHERE (((SPITR.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DATEBEGBAL = #" & txtDate.Text & "# " & _
                           "WHERE (((SPAREPARTS.SPAREPARTID)=" & flxBeg.TextMatrix(ActRow, 9) & "));"
            'Call Edit457(flxBeg.TextMatrix(ActRow, 10), txtDate.Text, ActCol)
            txtDate.Visible = False
            txtDate.Tag = ""
            Exit Sub
        End If
    End If
    flxBeg.TextMatrix(ActRow, ActCol) = txtDate.Text
    txtDate.Visible = False
    If CanAddRow And flxBeg.TextMatrix(ActRow, 9) = "" Then
        Call AddRecord
    End If
End Sub
                       
Private Sub UpdatePartNo(mySPID As Long)
    Dim rsSP As New ADODB.Recordset
    Dim rsTrans As New ADODB.Recordset
    Dim i As Integer
    
    i = ActRow
    With rsSP
        Qry = "SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & mySPID & "));"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            .Fields("PartNumber") = flxBeg.TextMatrix(i, 2)
            .Fields("Description") = flxBeg.TextMatrix(i, 3)
            .Fields("SupersedingPN") = flxBeg.TextMatrix(i, 10)
            .Fields("StorePrice") = IIf(IsNumeric(flxBeg.TextMatrix(i, 11)), flxBeg.TextMatrix(i, 11), 0)
            .Fields("Cost") = IIf(IsNumeric(flxBeg.TextMatrix(i, 12)), flxBeg.TextMatrix(i, 12), 0)
            .Fields("ApplicModel") = flxBeg.TextMatrix(i, 13)
            .Fields("Reference") = "BEG. BAL."
            .Fields("UpdatedBy") = UserID
            .Fields("LastUpdate") = Date
            .Update
        End If
        'flxBeg.TextMatrix(i, 9) = .Fields("SparePartID")
        .Close
    End With
End Sub

Private Sub AddRecord()
    Dim rsSP As New ADODB.Recordset
    Dim rsTrans As New ADODB.Recordset
    Dim i As Integer
    
    i = flxBeg.Rows - 1
    With rsSP
        .Open "SpareParts", LYRcon, adOpenKeyset, adLockOptimistic
        .AddNew
        .Fields("PartNumber") = flxBeg.TextMatrix(i, 2)
        .Fields("Description") = flxBeg.TextMatrix(i, 3)
        .Fields("SupersedingPN") = flxBeg.TextMatrix(i, 10)
        .Fields("StorePrice") = IIf(IsNumeric(flxBeg.TextMatrix(i, 11)), flxBeg.TextMatrix(i, 11), 0)
        .Fields("Cost") = IIf(IsNumeric(flxBeg.TextMatrix(i, 12)), flxBeg.TextMatrix(i, 12), 0)
        .Fields("QtyOnHand") = Val(flxBeg.TextMatrix(i, 1))
        .Fields("BegBal") = Val(flxBeg.TextMatrix(i, 1))
        .Fields("Location") = flxBeg.TextMatrix(i, 6)
        .Fields("ApplicModel") = flxBeg.TextMatrix(i, 13)
        .Fields("Remarks") = IIf(flxBeg.TextMatrix(i, 8) = "G", "GENUINE", "REPLACEMENT")
        .Fields("Reference") = "BEG. BAL."
        If IsDate(flxBeg.TextMatrix(i, 5)) Then .Fields("DateBegBal") = flxBeg.TextMatrix(i, 5)
        .Fields("UpdatedBy") = UserID
        .Fields("LastUpdate") = Date
        .Update
        flxBeg.TextMatrix(i, 9) = .Fields("SparePartID")
        .Close
    End With
    
    With rsTrans
        .Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
        .AddNew
        .Fields("Reference") = "BEG. BAL."
        .Fields("TransType") = "BEGBAL"
        .Fields("TransDate") = flxBeg.TextMatrix(i, 5)
        .Fields("SparePartID") = flxBeg.TextMatrix(i, 9)
        .Fields("Receive") = Val(flxBeg.TextMatrix(i, 1))
        .Fields("PackingList") = flxBeg.TextMatrix(i, 4)
        .Fields("Source") = flxBeg.TextMatrix(i, 7)
        .Update
        .Close
    End With
    flxBeg.AddItem ""
    If flxBeg.Rows > 18 Then flxBeg.Width = 11295
End Sub


Private Sub EditQty(MyPartID As Long, NewData As Integer)
    Dim rsSP As New ADODB.Recordset
    Dim rsTrans As New ADODB.Recordset
    Dim Qry As String
    
    On Error GoTo LocalErr
    Qry = "SELECT SPAREPARTS.QTYONHAND, SPAREPARTS.BEGBAL, SPAREPARTS.SPAREPARTID " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & MyPartID & "));"
    With rsSP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            .Fields("QtyOnHand") = NewData
            .Fields("BegBal") = NewData
            .Update
        End If
        .Close
    End With
    
    With rsTrans
        Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.SPAREPARTID)=" & MyPartID & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"

        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            .Fields("Receive") = NewData
            .Update
        End If
        .Close
    End With
    Exit Sub
LocalErr:
    MsgBox "There was a problem encountered in updating the record." & vbCrLf & "Please report the error message to the System Administrator." & vbCrLf & vbCrLf & _
    "Error Message: [" & CStr(Err.Number) & "] " & Err.Description, vbCritical + vbOKOnly, "Error"
    Resume Next
    
End Sub

Private Sub Edit457(MyPartID As Long, NewData As String, ColNo As Integer)
    Dim rsTrans As New ADODB.Recordset
    Dim Qry As String
    
    On Error GoTo LocalErr
    With rsTrans
        Qry = "SELECT SPITR.* From SPITR WHERE (((SPITR.SPAREPARTID)=" & MyPartID & ") AND ((SPITR.TRANSTYPE)='BEGBAL'));"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Select Case ColNo
                Case 4
                    .Fields("PackingList") = NewData
                Case 5
                    .Fields("TransDate") = NewData
                Case 7
                    .Fields("Source") = NewData
            End Select
            .Update
        End If
        .Close
    End With
    Exit Sub
LocalErr:
    MsgBox "There was a problem encountered in updating the record." & vbCrLf & "Please report the error message to the System Administrator." & vbCrLf & vbCrLf & _
    "Error Message: [" & CStr(Err.Number) & "] " & Err.Description, vbCritical + vbOKOnly, "Error"
    Resume Next
    
End Sub

Private Sub Edit36(MyPartID As Long, NewData As String, ColNo As Integer)
    Dim rsSP As New ADODB.Recordset
    Dim Qry As String
    
    
    On Error GoTo LocalErr
    Qry = "SELECT SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION, SPAREPARTS.PARTNUMBER, SPAREPARTS.SPAREPARTID " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & MyPartID & "));"
    With rsSP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If ColNo = 3 Then
                .Fields("Description") = NewData
            ElseIf ColNo = 6 Then
                .Fields("Location") = NewData
            ElseIf ColNo = 2 Then
                .Fields("PartNumber") = NewData
            End If
            .Update
        End If
        .Close
    End With
    Exit Sub
LocalErr:
    MsgBox "There was a problem encountered in updating the record." & vbCrLf & "Please report the error message to the System Administrator." & vbCrLf & vbCrLf & _
    "Error Message: [" & CStr(Err.Number) & "] " & Err.Description, vbCritical + vbOKOnly, "Error"
    Resume Next
    
End Sub

Private Sub DeleteSP(MyPartID As Long)
    Dim i As Integer
        
    LYRcon.Execute "DELETE SPAREPARTS.*, SPAREPARTS.SPAREPARTID From SPAREPARTS " & _
                   "WHERE (((SPAREPARTS.SPAREPARTID)=" & MyPartID & "));"
    LYRcon.Execute "DELETE SPITR.*, SPITR.TRANSTYPE, SPITR.SPAREPARTID From SPITR " & _
                   "WHERE (((SPITR.TRANSTYPE)='BEGBAL') AND ((SPITR.SPAREPARTID)=" & MyPartID & "));"
    cboFind.Clear
    For i = 1 To flxBeg.Rows - 1
        cboFind.AddItem flxBeg.TextMatrix(i, 2)
    Next i
End Sub

Private Sub SaveBegDesc(FldNo As Integer, NewData As String)
    Dim rsBeg As New ADODB.Recordset
    Dim Qry As String
    
    On Error GoTo LocalErr
    Qry = "SELECT BEGINDESC.PREPAREDBY, BEGINDESC.CHECKEDBY, BEGINDESC.INVDATE, BEGINDESC.LOCKED, BEGINDESC.UPDATEDBY, BEGINDESC.LASTUPDATE " & _
          "From BEGINDESC WHERE (((BEGINDESC.BEGCODE)='SP'));"
    With rsBeg
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If FldNo = 3 Then
                .Fields(FldNo) = True
            Else
                .Fields(FldNo) = NewData
            End If
            .Fields("UpdatedBy") = UserID
            .Fields("LastUpdate") = Date
            .Update
        End If
        .Update
        .Close
    End With
    Exit Sub
LocalErr:
    MsgBox "There was a problem encountered in updating." & vbCrLf & "Please report the error message to the System Administrator." & vbCrLf & vbCrLf & _
    "Error Message: [" & CStr(Err.Number) & "] " & Err.Description, vbCritical + vbOKOnly, "Error"
    Resume Next
    
End Sub

Private Sub txtPreparedBy_LostFocus()
    With txtPreparedBy
        If .Text <> "" And .Tag <> .Text Then Call SaveBegDesc(0, .Text)
    End With
End Sub

Private Sub txtSearch_Change()
    cmdNext.Tag = 1
End Sub

Private Sub txtSearch_Click()
    cmdNext.Tag = 1
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    
    If KeyCode = vbKeyEscape Then Call cmdCancel_Click
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        Tmp = SendMessage(txtSearch.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdNext_Click
End Sub
