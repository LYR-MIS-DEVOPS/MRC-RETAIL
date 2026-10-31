VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SPupdate 
   Caption         =   "Update Spare Parts"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   13935
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxITR 
      Height          =   1215
      Left            =   6000
      TabIndex        =   55
      Top             =   9120
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2143
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab tabUpdate 
      Height          =   5415
      Left            =   120
      TabIndex        =   34
      Tag             =   "XYF"
      Top             =   3480
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Show Spare Parts"
      TabPicture(0)   =   "SPupdate.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxSP"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdShow"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboShow"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Unused Spare Parts Archive"
      TabPicture(1)   =   "SPupdate.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FlxSrchArc"
      Tab(1).Control(1)=   "txtPartArc"
      Tab(1).Control(2)=   "txtFind"
      Tab(1).Control(3)=   "pBar"
      Tab(1).Control(4)=   "tabArc"
      Tab(1).Control(5)=   "flxArchive"
      Tab(1).Control(6)=   "cmdAddArc"
      Tab(1).Control(7)=   "cmdNxtArc"
      Tab(1).Control(8)=   "cmdCancelArc"
      Tab(1).Control(9)=   "cmdFind"
      Tab(1).Control(10)=   "lblArcNote"
      Tab(1).ControlCount=   11
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxSrchArc 
         Height          =   1695
         Left            =   -74520
         TabIndex        =   59
         Top             =   1560
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
      Begin VB.TextBox txtPartArc 
         Height          =   285
         Left            =   -69120
         TabIndex        =   58
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   -73560
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   485
         Width           =   2775
      End
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   105
         Left            =   -74880
         TabIndex        =   54
         Top             =   5280
         Visible         =   0   'False
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
      End
      Begin TabDlg.SSTab tabArc 
         Height          =   2295
         Left            =   -70200
         TabIndex        =   43
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4048
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Select Process"
         TabPicture(0)   =   "SPupdate.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblMDY"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdCancelTab"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdNextTab"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkAll"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkSpcfc"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtWithIN"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "updWin"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "updMDY"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         Begin MSComCtl2.UpDown updMDY 
            Height          =   285
            Left            =   3360
            TabIndex        =   50
            Top             =   600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            Max             =   4
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updWin 
            Height          =   285
            Left            =   2160
            TabIndex        =   48
            Top             =   600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtWithIN"
            BuddyDispid     =   196611
            OrigLeft        =   2160
            OrigTop         =   600
            OrigRight       =   2415
            OrigBottom      =   885
            Max             =   1000
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtWithIN 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   47
            Top             =   600
            Width           =   600
         End
         Begin VB.CheckBox chkSpcfc 
            Caption         =   "SPECIFIC PARTNUMBER"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   1080
            Width           =   3495
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "ALL PARTNUMBERS"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   2175
         End
         Begin lvButton.lvButtons_H cmdNextTab 
            Height          =   345
            Left            =   2160
            TabIndex        =   52
            Tag             =   "XYF"
            Top             =   1800
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            Caption         =   "Next"
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
            Image           =   "SPupdate.frx":0054
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCancelTab 
            Height          =   345
            Left            =   480
            TabIndex        =   53
            Tag             =   "XYF"
            Top             =   1800
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
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
            Image           =   "SPupdate.frx":05EE
            cBack           =   -2147483633
         End
         Begin VB.Label Label2 
            Caption         =   "Type all Part Numbers to be removed."
            Height          =   255
            Left            =   520
            TabIndex        =   51
            Top             =   1400
            Width           =   3255
         End
         Begin VB.Label lblMDY 
            Alignment       =   1  'Right Justify
            Caption         =   "Months"
            Height          =   255
            Left            =   2400
            TabIndex        =   49
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Unused within"
            Height          =   255
            Left            =   480
            TabIndex        =   46
            Top             =   630
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxArchive 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   38
         Tag             =   "XY"
         Top             =   840
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
         FormatString    =   "      |<PARTNUMBER              |< DESCRIPTION            |^ QTY.       |^DATE REMOVED |<SUPPLIER  |^LAST DATE USED|^DAYS UNUSED"
      End
      Begin VB.ComboBox cboShow 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   375
         Width           =   3135
      End
      Begin lvButton.lvButtons_H cmdShow 
         Height          =   345
         Left            =   3240
         TabIndex        =   35
         Tag             =   "XYF"
         ToolTipText     =   "Refresh Spare Parts"
         Top             =   360
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
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
         Image           =   "SPupdate.frx":0B88
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
         Height          =   4560
         Left            =   120
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   735
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   8043
         _Version        =   393216
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin lvButton.lvButtons_H cmdAddArc 
         Height          =   345
         Left            =   -63600
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   480
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         Caption         =   "Add SP to Archive"
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
         Image           =   "SPupdate.frx":1122
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNxtArc 
         Height          =   345
         Left            =   -63000
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   480
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         Caption         =   "Next"
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
         Image           =   "SPupdate.frx":16BC
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancelArc 
         Height          =   345
         Left            =   -64440
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   480
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
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
         Image           =   "SPupdate.frx":1C56
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   -74880
         TabIndex        =   57
         Tag             =   "XY"
         Top             =   440
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
         Image           =   "SPupdate.frx":21F0
         cBack           =   -2147483633
      End
      Begin VB.Label lblArcNote 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4320
      Top             =   3480
   End
   Begin VB.TextBox txtSuppCode 
      Height          =   285
      Left            =   10080
      TabIndex        =   33
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSrch 
      Height          =   1695
      Left            =   3000
      TabIndex        =   31
      Top             =   9120
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
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   15
      Tag             =   "XY"
      Top             =   720
      Width           =   13575
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         Left            =   8640
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   1406
         Width           =   4815
      End
      Begin VB.TextBox txtPartNo 
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
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   263
         Width           =   3135
      End
      Begin VB.TextBox txtSPid 
         Height          =   285
         Left            =   7680
         TabIndex        =   30
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboType 
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
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   1015
         Width           =   3135
      End
      Begin VB.TextBox txtApplic 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   8640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   1782
         Width           =   4815
      End
      Begin VB.TextBox txtLoc 
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
         Left            =   8640
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   1023
         Width           =   4815
      End
      Begin VB.TextBox txtDesc2 
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
         Left            =   8640
         MaxLength       =   150
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   639
         Width           =   4815
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   8640
         MaxLength       =   150
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   263
         Width           =   4815
      End
      Begin VB.TextBox txtQty 
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
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtMax 
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
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   1782
         Width           =   3135
      End
      Begin VB.TextBox txtReorder 
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
         MaxLength       =   5
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   1406
         Width           =   3135
      End
      Begin VB.TextBox txtSuper 
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
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   639
         Width           =   3135
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "SUPPLIER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   6720
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   1455
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "APPLICABLE MODELS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   6720
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   1785
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "LOCATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   6720
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "OTHER DESCRIPTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   6720
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   690
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "DESCRIPTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   6720
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "QTY. ON-HAND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   135
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   2197
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "MAXIMUM ALLOW."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   135
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   1834
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "REORDER POINT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   135
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   1458
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   1075
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "SUPERSEDING P.N."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   691
         Width           =   1800
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Caption         =   "PART NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   315
         Width           =   1800
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   11
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPupdate.frx":278A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Update"
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
      Image           =   "SPupdate.frx":2D24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPupdate.frx":32BE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   14
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
      Image           =   "SPupdate.frx":3858
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14055
      m_MDI_Height    =   9450
      OriginalBorder  =   2
      OriginalCaption =   27
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE SPAREPARTS"
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
      Left            =   10680
      TabIndex        =   18
      Top             =   480
      Width           =   3000
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE SPAREPARTS"
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
      Left            =   10710
      TabIndex        =   19
      Top             =   510
      Width           =   3000
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
      Left            =   210
      TabIndex        =   16
      Top             =   440
      Width           =   3000
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
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Width           =   3000
   End
   Begin VB.Menu mnuDelB4Archive 
      Caption         =   "Remove in Flx Before final"
      Visible         =   0   'False
      Begin VB.Menu mnuRemove 
         Caption         =   "Do Not Include in Archive"
      End
   End
   Begin VB.Menu mnuRollBack 
      Caption         =   "RollBack"
      Visible         =   0   'False
      Begin VB.Menu mnuRollBckPN 
         Caption         =   "Rollback Part Number"
      End
   End
   Begin VB.Menu mnuFlxArchive 
      Caption         =   "Archive Rows"
      Visible         =   0   'False
      Begin VB.Menu mnuArcADD 
         Caption         =   "Add New Row"
      End
      Begin VB.Menu mnuArcDel 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "SPupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsSP As New ADODB.Recordset
Dim rsSupplier As New ADODB.Recordset
Dim rsPrtsArc As New ADODB.Recordset
Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean, SuppSrch As Boolean, ComboSrch As Boolean
Dim ActRow As Integer, ActCol As Integer, FullAccess As Boolean, fRow As Integer, fCol As Integer

Private Sub ClearEntries()
    txtPartNo.Text = ""
    txtSuper.Text = ""
    cboType.Text = ""
    txtReorder.Text = ""
    txtMax.Text = ""
    txtQty.Text = ""
    txtDesc.Text = ""
    txtDesc2.Text = ""
    txtLoc.Text = ""
    txtApplic.Text = ""
    txtSPID.Text = ""
    cboSupplier.Text = ""
    txtSuppCode.Text = ""
End Sub

Private Sub LockData(bln As Boolean)
    txtPartNo.Locked = Not bln
    txtSuper.Locked = bln
    cboType.Locked = bln 'IIf(cboType.Text <> "", True, Bln)
    txtReorder.Locked = bln
    txtMax.Locked = bln
    txtDesc.Locked = bln
    txtDesc2.Locked = bln
    txtLoc.Locked = IIf(txtLoc.Text <> "", True, bln)
    txtApplic.Locked = bln
    txtSPID.Locked = bln
    cboSupplier.Locked = bln
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
            txtSuppCode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            'chkAllSupp.Value = 0
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppCode.Text = ""
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
    If KeyAscii = 13 Then txtApplic.SetFocus
End Sub

Private Sub cbotype_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If ComboSrch Then
        ComboSrch = False
        Exit Sub
    End If
    With cboType
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

Private Sub cboType_Click()
    Call cbotype_Change
End Sub

Private Sub cboType_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then ComboSrch = True
End Sub
Private Sub cboType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub chkAll_Click()
    If chkAll.Value = 1 Then chkSpcfc.Value = 0
End Sub

Private Sub chkSpcfc_Click()
    If chkSpcfc.Value = 1 Then chkAll.Value = 0
End Sub

Private Sub cmdAddArc_Click()
    
    If Not FullAccess Then
        MsgBox "Sorry, you do not have full access to this transaction!", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
        
    tabArc.Top = flxArchive.Top + 1000
    tabArc.Left = (tabUpdate.Width - tabArc.Width) / 2
    tabArc.Visible = True
    cmdFind.Visible = False
    txtFind.Visible = False
    Frame1.Enabled = False
    cmdTrans(0).Enabled = False
    cmdTrans(1).Enabled = False
    tabUpdate.TabEnabled(0) = False
    cmdAddArc.Visible = False
    lblArcNote.Caption = ""
End Sub

Private Sub cmdCancelArc_Click()
    Dim rsArchive As New ADODB.Recordset
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Cancel moving of Spare Parts to archive?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    Frame1.Enabled = True
    cmdTrans(0).Enabled = True
    cmdTrans(1).Enabled = True
    tabUpdate.TabEnabled(0) = True
    cmdAddArc.Visible = True
    cmdFind.Visible = True
    txtFind.Visible = True
    txtPartArc.Visible = False
    FlxSrchArc.Visible = False
    cmdCancelArc.Visible = False
    cmdNxtArc.Visible = False
    lblArcNote.Caption = ""
    'Display records in Archive
    Qry = "SELECT ARCHIVESP.PARTNUMBER, ARCHIVESP.DESCRIPTION, ARCHIVESP.QTYONHAND, ARCHIVESP.DESCRIPTION2, ARCHIVESP.SPAREPARTID FROM ARCHIVESP ORDER BY PARTNUMBER;"
    
    flxArchive.Rows = 1
    With rsArchive
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Desc2 = IIf(IsNull(!Description2), "", Split(!Description2, "|"))
                flxArchive.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & Desc2(0) & vbTab & Desc2(1) & vbTab & Desc2(2) & vbTab & Desc2(3)
                .MoveNext
            Loop
        End If
        .Close
    End With
    
End Sub

Private Sub cmdCancelTab_Click()
    Dim rsArchive As New ADODB.Recordset
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Cancel moving of Spare Parts to archive?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    tabArc.Visible = False
    cmdFind.Visible = True
    txtFind.Visible = True
    Frame1.Enabled = True
    cmdTrans(0).Enabled = True
    cmdTrans(1).Enabled = True
    tabUpdate.TabEnabled(0) = True
    cmdAddArc.Visible = True
    'Display records in Archive
    Qry = "SELECT ARCHIVESP.PARTNUMBER, ARCHIVESP.DESCRIPTION, ARCHIVESP.QTYONHAND, ARCHIVESP.DESCRIPTION2, ARCHIVESP.SPAREPARTID FROM ARCHIVESP ORDER BY PARTNUMBER;"
    With rsArchive
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxArchive.Rows = 1
            .MoveFirst
            Do While Not .EOF
                Desc2 = IIf(IsNull(!Description2), "", Split(!Description2, "|"))
                flxArchive.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & Desc2(0) & vbTab & Desc2(1) & vbTab & Desc2(2) & vbTab & Desc2(3)
                .MoveNext
            Loop
        End If
        .Close
    End With
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxArchive
        For i = fRow To .Rows - 1
            For j = fCol To 7
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 7: .RowSel = i
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

Private Sub cmdNextTab_Click()
    Dim rsArc As New ADODB.Recordset
    Dim Qry As String, MySup As String, UDay As String, LastDate As Date
    
    'validate txtwithin
    If chkAll.Value = 0 And chkSpcfc.Value = 0 Then
        MsgBox "Please select type of Process in removing Partnumber", vbExclamation + vbOKOnly, "Move to Archive"
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        If Not IsNumeric(txtWithIN.Text) Then
            MsgBox "Invalid value of unsused number of " & lblMDY.Caption & "?", vbExclamation + vbOKOnly, "Invalid"
            Exit Sub
        ElseIf Val(txtWithIN.Text > 1000) Then
                MsgBox "Invalid value of unsused number of " & lblMDY.Caption & "?", vbExclamation + vbOKOnly, "Invalid"
                Exit Sub
        End If
        If InStr(lblMDY.Caption, "Month") > 0 Then
            LastDate = DateAdd("m", Val(txtWithIN.Text) * -1, Date)
        ElseIf InStr(lblMDY.Caption, "Day") > 0 Then
            LastDate = DateAdd("d", Val(txtWithIN.Text) * -1, Date)
        ElseIf InStr(lblMDY.Caption, "Month") > 0 Then
            LastDate = DateAdd("y", Val(txtWithIN.Text) * -1, Date)
        End If
        tabArc.Enabled = False
        '***Access
        'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.SPAREPARTID, SPAREPARTS.SUPPLIERCODE " & _
              "From SPAREPARTS Where (((SPAREPARTS.QTYONHAND) = 0)) ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.SPAREPARTID, SPAREPARTS.SUPPLIERCODE " & _
              "From SPAREPARTS Where SPAREPARTS.QTYONHAND = 0 ORDER BY SPAREPARTS.PARTNUMBER;"
        
        pBar.Top = flxArchive.Top + flxArchive.Height
        pBar.Left = flxArchive.Left
        pBar.Width = flxArchive.Width
        pBar.Value = 0
        flxArchive.Rows = 1
        With rsArc
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                pBar.Visible = True
                pBar.Max = .RecordCount
                Do While Not .EOF
                '***Access
                'Qry = "SELECT SPITR.TRANSDATE From SPITR " & _
                      "WHERE (((SPITR.SPAREPARTID)=" & !SparepartID & ") AND ((SPITR.TRANSTYPE)<>'CANCELLED' And (SPITR.TRANSTYPE)<>'ORDER TEMP')) ORDER BY SPITR.TRANSDATE DESC;"
                'mySQL
                Qry = "SELECT SPITR.TRANSDATE From SPITR " & _
                      "WHERE SPITR.SPAREPARTID=" & !SparepartID & " AND SPITR.TRANSTYPE<>'CANCELLED' And SPITR.TRANSTYPE<>'ORDER TEMP' ORDER BY SPITR.TRANSDATE DESC;"
             
                UDay = GetFieldValue(Qry)
                If IsDate(UDay) Then
                    If LastDate > DateValue(UDay) Then
                        MySup = IIf(IsNull(!SUPPLIERCODE), "", !SUPPLIERCODE)
                        '***Access
                        'Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & MySup & "'));"
                        
                        'mySQL
                        Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & MySup & "';"
                        
                        MySup = GetFieldValue(Qry)
                        flxArchive.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & Format(Date, "mm/dd/yyyy") & vbTab & MySup & vbTab & UDay & vbTab & DateDiff("D", Date, DateValue(UDay)) & vbTab & !SparepartID
                    End If
                Else
                    MySup = IIf(IsNull(!SUPPLIERCODE), "", !SUPPLIERCODE)
                    
                    '***Access
                    'Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & MySup & "'));"
                    'mySQL
                    Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & MySup & "';"
                    
                    MySup = GetFieldValue(Qry)
                    flxArchive.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & Format(Date, "mm/dd/yyyy") & vbTab & MySup & vbTab & "Unused" & vbTab & "Unknown" & vbTab & !SparepartID
                End If
                .MoveNext
                pBar.Value = pBar.Value + 1
                DoEvents
                Loop
            Else
                .Close
                MsgBox "No available spareparts for your criteria!", vbExclamation + vbOKOnly, "0 records found"
                Exit Sub
            End If
            .Close
        End With
        tabArc.Enabled = True
        pBar.Visible = False
        If flxArchive.Rows > 1 Then
            tabArc.Visible = False
            cmdCancelArc.Visible = True
            cmdNxtArc.Visible = True
            lblArcNote.Caption = "Click NEXT to move all records in grid to Unused Part Number Archive"
            tabArc.Tag = "ALL PARTS"
        End If
    Else    'Specific Partnumbers
        tabArc.Visible = False
        cmdCancelArc.Visible = True
        cmdNxtArc.Visible = True
        lblArcNote.Caption = "Type Part Number on grid and Click NEXT to move records in grid to Unused Part Number Archive"
        tabArc.Tag = "SPECIFIC PARTS"
        flxArchive.Rows = 1
        flxArchive.Rows = 2
        If rsPrtsArc.State = adStateOpen Then rsPrtsArc.Close
        '***Access
        'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE " & _
              "From SPAREPARTS Where (((SPAREPARTS.QTYONHAND) < 1)) ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE " & _
              "From SPAREPARTS Where SPAREPARTS.QTYONHAND < 1 ORDER BY SPAREPARTS.PARTNUMBER;"
        
        rsPrtsArc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Set FlxSrchArc.DataSource = rsPrtsArc
        With FlxSrchArc
            .ColWidth(0) = 0
            .ColWidth(1) = 0
            .ColWidth(2) = flxArchive.ColWidth(1)
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColWidth(5) = 0
        End With
    End If
End Sub

Private Sub cmdNxtArc_Click()
    Dim i As Integer, MyDesc As String
    
    With flxArchive
        If chkAll.Value = 0 Then
            FlxSrchArc.Visible = False
            txtPartArc.Visible = False
            For i = .Rows - 1 To 1 Step -1
                If .TextMatrix(i, 8) = "" Then .RemoveItem i
            Next i
        End If
        pBar.Value = 0
        pBar.Max = .Rows - 1
        pBar.Visible = True
        lblArcNote.Caption = "Moving records to Archive..."
        For i = 1 To .Rows - 1
            '***Access
            'LYRcon.Execute "INSERT INTO archivesp SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            'mySQL
            LYRcon.Execute "INSERT INTO archivesp SELECT SPAREPARTS.* From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            'Get SPITR
            '***Access
            'LYRcon.Execute "INSERT INTO ARCHIVESPITR SELECT SPITR.* From SPITR WHERE (((SPITR.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            'mySQL
            LYRcon.Execute "INSERT INTO ARCHIVESPITR SELECT SPITR.* From SPITR WHERE SPITR.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            MyDesc = .TextMatrix(i, 4) & "|" & .TextMatrix(i, 5) & "|" & .TextMatrix(i, 6) & "|" & .TextMatrix(i, 7)
            '***Access
            'LYRcon.Execute "UPDATE ARCHIVESP SET ARCHIVESP.DESCRIPTION2 = '" & MyDesc & "' WHERE (((ARCHIVESP.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            'mySQL
            LYRcon.Execute "UPDATE ARCHIVESP SET ARCHIVESP.DESCRIPTION2 = '" & MyDesc & "' WHERE ARCHIVESP.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            '***Access
            'LYRcon.Execute "DELETE SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 8) & "));"
            'mySQL
            LYRcon.Execute "DELETE From SPAREPARTS WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 8) & ";"
            pBar.Value = pBar.Value + 1
            DoEvents
        Next i
        lblArcNote.Caption = ""
        pBar.Visible = False
        MsgBox "Part Number successfully moved to Archive!", vbInformation + vbOKOnly, "Move Complete"
        Frame1.Enabled = True
        cmdTrans(0).Enabled = True
        cmdTrans(1).Enabled = True
        tabUpdate.TabEnabled(0) = True
        cmdAddArc.Visible = True
        cmdCancelArc.Visible = False
        cmdNxtArc.Visible = False
        tabArc.Tag = ""
    End With
            
End Sub

Private Sub cmdShow_Click()
    If cboShow.ListIndex = -1 Then
        MsgBox "Please select an item from the list to continue!", vbExclamation + vbOKOnly, "View Spare Parts"
        Exit Sub
    End If
    Timer1.Enabled = True
    Call RXResizer1_ResizeComplete
End Sub

Private Sub BtnSaveCancel(bln As Boolean)
    If bln Then
        cmdTrans(0).Visible = False
        cmdTrans(1).Visible = False
        'cmdTrans(4).Visible = False
        cmdTrans(2).Left = 0
        cmdTrans(2).Visible = True
        cmdTrans(3).Left = cmdTrans(2).Width
        cmdTrans(3).Visible = True
    Else
        cmdTrans(2).Visible = False
        cmdTrans(3).Visible = False
        cmdTrans(0).Visible = True
        cmdTrans(1).Visible = True
        'cmdTrans(4).Visible = True
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer, ErrSec As String, Qry As String
    Dim Tmp() As String
    
    'On Error GoTo LocalErr
    Select Case Index
        Case 0  'Update
            If Not FullAccess Then
                MsgBox "Sorry, you do not have full access to this transaction!", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtSPID.Text = "" Then
                MsgBox "No Valid Spare Parts available for edit!", vbExclamation + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Call LockData(False)
            Call BtnSaveCancel(True)
            cmdTrans(3).Tag = txtSuper.Text & "|" & cboType.Text & "|" & txtReorder.Text & "|" & txtMax.Text & "|" & txtDesc2.Text & "|" & txtLoc.Text & "|" & txtApplic.Text
            txtPartNo.Enabled = False
            txtQty.Enabled = False
            'txtDesc.Enabled = False
            flxSP.Enabled = False
            cboShow.Enabled = False
            cmdShow.Enabled = False
            txtSuper.SetFocus
        Case 1  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
        Case 2  'Save
            'If cboType.ListIndex = -1 Then
            '    MsgBox "Please select Type of Spare Parts from list.", vbExclamation + vbOKOnly, "Validation Error"
            '    Exit Sub
            'End If
            Ans = MsgBox("Do you want to Save Changes?", vbYesNo + vbQuestion, "Save")
            If Ans = vbNo Then Exit Sub
            'Check if location is used my other partnumber
            ErrSec = SPCrntLoc
            If ErrSec <> "" Then
                Ans = MsgBox(ErrSec & vbCrLf & "Do you want to Continue Saving?", vbQuestion + vbYesNo, "Location Already in Used!")
                If Ans = vbNo Then Exit Sub
            End If
            'Save changes
            ErrSec = "Updating Table Spare Parts"
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & txtDesc.Text & "', SPAREPARTS.DESCRIPTION2 = '" & txtDesc2.Text & "', SPAREPARTS.LOCATION = '" & txtLoc.Text & "', SPAREPARTS.APPLICMODEL = '" & txtApplic.Text & "', SPAREPARTS.SUPERSEDINGPN = '" & txtSuper.Text & "', SPAREPARTS.REMARKS = '" & cboType.Text & "', SPAREPARTS.REORDERPOINT = " & Val(txtReorder.Text) & ", SPAREPARTS.SUPPLIERCODE = '" & txtSuppCode.Text & "', SPAREPARTS.MAXALLOW = " & Val(txtMax.Text) & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & txtSPid.Text & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION = '" & txtDesc.Text & "', SPAREPARTS.DESCRIPTION2 = '" & txtDesc2.Text & "', SPAREPARTS.LOCATION = '" & txtLoc.Text & "', SPAREPARTS.APPLICMODEL = '" & txtApplic.Text & "', SPAREPARTS.SUPERSEDINGPN = '" & txtSuper.Text & "', SPAREPARTS.REMARKS = '" & cboType.Text & "', SPAREPARTS.REORDERPOINT = " & Val(txtReorder.Text) & ", SPAREPARTS.SUPPLIERCODE = '" & txtSuppCode.Text & "', SPAREPARTS.MAXALLOW = " & Val(txtMax.Text) & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & txtSPID.Text & ";"
            '***Access
            'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & txtDesc.Text & "',  SOURCESP.APPLICMODEL = '" & txtApplic.Text & "', SOURCESP.SUPERSEDING = '" & txtSuper.Text & "' " & _
                          " WHERE (((SOURCESP.PARTNUMBER)='" & txtPartNo.Text & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.DESCRIPTION = '" & txtDesc.Text & "',  SOURCESP.APPLICMODEL = '" & txtApplic.Text & "', SOURCESP.SUPERSEDING = '" & txtSuper.Text & "' " & _
                          " WHERE SOURCESP.PARTNUMBER='" & txtPartNo.Text & "';"
            If ErrSec <> "" Then    'Update flxSP with changes if no error occured
                flxSP.TextMatrix(ActRow, 2) = txtDesc.Text
                flxSP.TextMatrix(ActRow, 3) = txtDesc2.Text
                flxSP.TextMatrix(ActRow, 4) = txtLoc.Text
                flxSP.TextMatrix(ActRow, 5) = txtApplic.Text
                flxSP.TextMatrix(ActRow, 6) = cboType.Text
                flxSP.TextMatrix(ActRow, 7) = txtSuper.Text
                flxSP.TextMatrix(ActRow, 9) = txtReorder.Text
                flxSP.TextMatrix(ActRow, 10) = txtMax.Text
                For i = 1 To flxSrch.Rows - 1
                    If flxSrch.TextMatrix(i, 1) = txtSPID.Text Then
                        flxSrch.TextMatrix(i, 4) = txtDesc2.Text
                        flxSrch.TextMatrix(i, 5) = txtLoc.Text
                        flxSrch.TextMatrix(i, 6) = txtApplic.Text
                        flxSrch.TextMatrix(i, 8) = cboType.Text
                        flxSrch.TextMatrix(i, 7) = txtSuper.Text
                        flxSrch.TextMatrix(i, 9) = txtReorder.Text
                        flxSrch.TextMatrix(i, 10) = txtMax.Text
                        Exit For
                    End If
                Next i
            End If
                
            MsgBox "Spare Parts successfully updated!", vbInformation + vbOKOnly, "Save"
            Call LockData(True)
            Call BtnSaveCancel(False)
            txtPartNo.Enabled = True
            txtQty.Enabled = True
            txtDesc.Enabled = True
            flxSP.Enabled = True
            cboShow.Enabled = True
            cmdShow.Enabled = True
            'txtPartNo.SetFocus
        
        Case 3  'Cancel
            Ans = MsgBox("Do you want to Cancel Changes?", vbYesNo + vbQuestion, "Save")
            If Ans = vbNo Then Exit Sub
            Tmp = Split(cmdTrans(3).Tag, "|")
            txtSuper.Text = Tmp(0)
            cboType.Text = Tmp(1)
            txtReorder.Text = Tmp(2)
            txtMax.Text = Tmp(3)
            txtDesc2.Text = Tmp(4)
            txtLoc.Text = Tmp(5)
            txtApplic.Text = Tmp(6)
            Call LockData(True)
            Call BtnSaveCancel(False)
            txtPartNo.Enabled = True
            txtQty.Enabled = True
            txtDesc.Enabled = True
            flxSP.Enabled = True
            cboShow.Enabled = True
            cmdShow.Enabled = True
            txtPartNo.SetFocus
    End Select
    Exit Sub
LocalErr:
    MsgBox "Error " & ErrSec & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
    ErrSec = ""
    Resume Next
End Sub

Function SPCrntLoc() As String
    Dim rsLoc As New ADODB.Recordset
    Dim Qry As String
    
    If txtLoc.Text = "" Then SPCrntLoc = "": Exit Function
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.LOCATION From SPAREPARTS " & _
          "WHERE (((SPAREPARTS.LOCATION) Like '%" & txtLoc.Text & "%'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.LOCATION From SPAREPARTS " & _
          "WHERE SPAREPARTS.LOCATION Like '%" & txtLoc.Text & "%';"
    
    With rsLoc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            SPCrntLoc = !PartNumber & " uses the Location " & !Location
        Else
            SPCrntLoc = ""
        End If
        .Close
    End With
End Function

Private Sub flxArchive_Click()
    If flxArchive.Col = 1 And tabArc.Tag = "SPECIFIC PARTS" Then
        ActRow = flxArchive.Row
        ActCol = flxArchive.Col
        With txtPartArc
            If flxArchive.TextMatrix(ActRow, 1) <> "" Then .Tag = flxArchive.TextMatrix(ActRow, 1) & "|" & flxArchive.TextMatrix(ActRow, 2) & "|" & flxArchive.TextMatrix(ActRow, 3) & "|" & flxArchive.TextMatrix(ActRow, 4) & "|" & flxArchive.TextMatrix(ActRow, 5) & "|" & flxArchive.TextMatrix(ActRow, 6) & "|" & flxArchive.TextMatrix(ActRow, 7) & "|" & flxArchive.TextMatrix(ActRow, 8)
            .Top = flxArchive.CellTop + flxArchive.Top
            .Left = flxArchive.CellLeft + flxArchive.Left
            .Width = flxArchive.CellWidth
            .Text = flxArchive.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End If
End Sub

Private Sub flxArchive_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If flxArchive.Col = 1 And tabArc.Tag = "ALL PARTS" Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuDelB4Archive)
        End If
    ElseIf flxArchive.Col = 0 And tabArc.Tag = "SPECIFIC PARTS" Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuFlxArchive)
        End If
    ElseIf flxArchive.Col = 1 And tabArc.Tag = "" And flxArchive.TextMatrix(flxArchive.Row, 1) <> "" Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuRollBack)
        End If
    End If
End Sub

Private Sub flxSP_EnterCell()
    If rsSP.State = adStateClosed Then Exit Sub
    If rsSP.RecordCount = 0 Then Exit Sub
    Call ClearEntries
    With flxSP
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 1)
        txtSuper.Text = .TextMatrix(.Row, 7)
        cboType.Text = .TextMatrix(.Row, 6)
        txtReorder.Text = .TextMatrix(.Row, 9)
        txtMax.Text = .TextMatrix(.Row, 10)
        txtQty.Text = .TextMatrix(.Row, 8)
        txtDesc.Text = .TextMatrix(.Row, 2)
        txtDesc2.Text = .TextMatrix(.Row, 3)
        txtLoc.Text = .TextMatrix(.Row, 4)
        txtApplic.Text = .TextMatrix(.Row, 5)
        txtSPID.Text = .TextMatrix(.Row, 11)
        ActRow = .Row
    End With
End Sub

Private Sub flxSrch_Click()
    FromFlx = True
End Sub

Private Sub flxSrch_EnterCell()
    With flxSrch
        FromFlx = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub flxSrch_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        flxSrch.Visible = False
        txtPartNo.SetFocus
    End If
End Sub
                                                 
Private Sub flxSrch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
    End If
End Sub

Private Sub FlxSrchArc_EnterCell()
    txtPartArc.Text = FlxSrchArc.TextMatrix(FlxSrchArc.Row, 2)
End Sub

Private Sub FlxSrchArc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call txtPartArc_LostFocus
End Sub

Private Sub Form_Load()
    Dim rsPType As New ADODB.Recordset
    Dim Qry As String
        
    
    '***Access
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=36));"
    
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=36;"
    FullAccess = IIf(GetFieldValue(Qry) = "FULL ACCESS", True, False)
    
    Qry = "SELECT DISTINCT spareparts.REMARKS From spareparts Order By spareparts.REMARKS;"
    With rsPType
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            cboType.Clear
            .MoveFirst
            Do While Not .EOF
                cboType.AddItem !Remarks
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    cboShow.AddItem "SHOW ALL"
    cboShow.AddItem "BLANK APPLIC. MODEL"
    cboShow.AddItem "BLANK LOCATION"
    cboShow.AddItem "BLANK OTHER DESCRIPTION"
    cboShow.AddItem "BLANK TYPE"
    cboShow.AddItem "BLANK SUPERSEDING P.N."
    
    cboShow.ListIndex = 0
    cboShow.Text = cboShow.List(0)
    
    Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
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
    Timer1.Enabled = True
    tabUpdate.Tab = 0
    Call LockData(True)
    MotoRep.sBar.Panels(6).Text = "Updating Spare Parts"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsParts.State = adStateOpen Then rsParts.Close
    If rsSP.State = adStateOpen Then rsSP.Close
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub Form_Resize()
    txtPartNo.Tag = "XYF"   'expected mo resize ang textbox during form resize
End Sub

Private Sub mnuArcADD_Click()
    With flxArchive
        If .TextMatrix(.Rows - 1, 8) = "" Then
            MsgBox "Please enter a correct Part Number on the last row!", vbExclamation + vbOKOnly, "Add row"
        Else
            .AddItem ""
        End If
    End With
End Sub

Private Sub mnuArcDel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete row?", vbQuestion + vbYesNo, "Remove")
    If Ans = vbYes Then
        If flxArchive.Rows = 2 Then
            flxArchive.Rows = 1
            flxArchive.Rows = 2
        Else
            flxArchive.RemoveItem flxArchive.Row
        End If
    End If
End Sub

Private Sub mnuRemove_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to remove Part Number from list?", vbQuestion + vbYesNo, "Remove")
    If Ans = vbYes Then
        If flxArchive.Rows = 2 Then
            flxArchive.Rows = 1
            flxArchive.Rows = 2
        Else
            flxArchive.RemoveItem flxArchive.Row
        End If
    End If
End Sub

Private Sub mnuRollBckPN_Click()
    Dim rsArchITR As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, myPN As String
    Dim i As Integer, Ans As Integer, MyID As Long
            
    If Not FullAccess Then
        MsgBox "Sorry, you do not have full access to this transaction!", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    'Validate if PartNumber already exist
    myPN = flxArchive.TextMatrix(flxArchive.Row, 1)
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & myPN & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & myPN & "';"
    
    If CtrlNumExist(Qry) Then
        MsgBox "Cannot Rollback Part Number " & flxArchive.TextMatrix(flxArchive.Row, 1) & " because it already exist!", vbCritical + vbOKOnly, "Record Exist"
        Exit Sub
    End If
    '2183735
    Ans = MsgBox("Do you want to Rollback Part Number:  " & flxArchive.TextMatrix(flxArchive.Row, 1) & "?", vbQuestion + vbYesNo, "Rollback")
    If Ans = vbNo Then Exit Sub
    MyID = Val(flxArchive.TextMatrix(flxArchive.Row, 8))
    '***Access
    'LYRcon.Execute "INSERT INTO SPAREPARTS SELECT ARCHIVESP.* From ARCHIVESP WHERE (((ARCHIVESP.SPAREPARTID)=" & MyId & "));"
    'mySQL
    LYRcon.Execute "INSERT INTO SPAREPARTS SELECT ARCHIVESP.* From ARCHIVESP WHERE ARCHIVESP.SPAREPARTID=" & MyID & ";"
    
    'Update SPITR
    '***Access
    'Qry = "SELECT ARCHIVESPITR.* From ARCHIVESPITR WHERE (((ARCHIVESPITR.SPAREPARTID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT ARCHIVESPITR.* From ARCHIVESPITR WHERE ARCHIVESPITR.SPAREPARTID=" & MyID & ";"
    
    rsArchITR.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxITR.DataSource = rsArchITR
    flxITR.Refresh
    If rsArchITR.RecordCount > 0 Then
        '***Access
        'Qry = "SELECT SPAREPARTS.SPAREPARTID From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & myPN & "'));"
        'Qry = "SELECT SPAREPARTS.SPAREPARTID From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & myPN & "';"
        'MyId = Val(GetFieldValue(Qry))
        'With flxITR
        '    rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
        '    For i = 1 To .Rows - 1
        '        rsITR.AddNew
        '        rsITR!REFERENCE = .TextMatrix(.Row, 1)
        '        rsITR!TRANSTYPE = .TextMatrix(.Row, 2)
        '        rsITR!TRANSDATE = .TextMatrix(.Row, 3)
        '        rsITR!SparePartID = MyId
        '        rsITR!RECEIVE = Val(.TextMatrix(.Row, 5))
        '        rsITR!OUT = Val(.TextMatrix(.Row, 6))
        '        rsITR!Order = Val(.TextMatrix(.Row, 7))
        '        rsITR!BALANCE = Val(.TextMatrix(.Row, 8))
        '        rsITR!ORDERBAL = .TextMatrix(.Row, 9)
        '        rsITR!Price = Val(.TextMatrix(.Row, 10))
        '        rsITR!Remarks = .TextMatrix(.Row, 11)
        '        rsITR!PACKINGLIST = .TextMatrix(.Row, 12)
        '        rsITR!Source = .TextMatrix(.Row, 13)
        '        rsITR!UPDATEDBY = Val(.TextMatrix(.Row, 14))
        '        rsITR!LASTUPDATE = .TextMatrix(.Row, 15)
        '        rsITR.Update
        '    Next i
        '    rsITR.Close
        'End With
        
        'mySQL Statement
        Qry = "SELECT SPAREPARTS.SPAREPARTID From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & myPN & "';"
        MyID = Val(GetFieldValue(Qry))
        With flxITR
            For i = 1 To .Rows - 1
                LYRcon.Execute "INSERT INTO spitr(REFERENCE, TRANSTYPE, TRANSDATE, SPAREPARTID, RECEIVE, `OUT`," & _
                    "`ORDER`, BALANCE, ORDERBAL, PRICE, REMARKS, PACKINGLIST, SOURCE, UPDATEDBY, LASTUPDATE)" & _
                    "VALUES('" & .TextMatrix(.Row, 1) & "', '" & .TextMatrix(.Row, 2) & "', '" & Format(.TextMatrix(.Row, 3), "yyyy-mm-dd") & "', " & _
                    "" & MyID & ", " & Val(.TextMatrix(.Row, 5)) & ", " & Val(.TextMatrix(.Row, 6)) & ", " & _
                    "" & Val(.TextMatrix(.Row, 7)) & ", " & Val(.TextMatrix(.Row, 8)) & ", " & .TextMatrix(.Row, 9) & ", " & _
                    "" & Val(.TextMatrix(.Row, 10)) & ", '" & .TextMatrix(.Row, 11) & "', '" & .TextMatrix(.Row, 12) & "', " & _
                    "'" & .TextMatrix(.Row, 13) & "', " & Val(.TextMatrix(.Row, 14)) & ", '" & Format(.TextMatrix(.Row, 15), "yyyy-mm-dd hh:mm:ss") & "');"
            Next i
        End With
    
    End If
    rsArchITR.Close
    MsgBox "Part Number " & myPN & " successfully activated!", vbInformation + vbOKOnly, "Rollback Complete"
    flxArchive.RemoveItem flxArchive.Row
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxSP
        .FormatString = "   |<PART NUMBER|<DESCRIPTION|<OTHER DESCRIPTION|<LOCATION|<APPLIC. MODEL |<TYPE |<SUPERSEDING PN|>QTY. ON-HAND|>REORDER|>MAX ALLOW."
        .ColWidth(0) = 300
        .ColWidth(1) = .Width * 0.15
        .ColWidth(2) = .Width * 0.2
        .ColWidth(3) = .Width * 0.15
        .ColWidth(4) = .Width * 0.15
        .ColWidth(5) = .Width * 0.15
        .ColWidth(6) = .Width * 0.15
        .ColWidth(7) = .Width * 0.15
        .ColWidth(8) = .Width * 0.1
        .ColWidth(9) = .Width * 0.1
        .ColWidth(10) = .Width * 0.1
        .ColWidth(11) = 0
        lblHeading(2).Left = (.Left + .Width) - lblHeading(2).Width
        lblHeading(3).Left = (.Left + .Width) - lblHeading(3).Width + 30
    
    End With
    With flxArchive
        .ColWidth(0) = 300
        .ColWidth(1) = .Width * 0.18
        .ColWidth(2) = .Width * 0.3
        .ColWidth(3) = .Width * 0.07
        .ColWidth(4) = .Width * 0.1
        .ColWidth(5) = .Width * 0.13
        .ColWidth(6) = .Width * 0.1
        .ColWidth(7) = .Width * 0.1
        .ColWidth(8) = 0
    End With
    
End Sub

Private Sub Timer1_Timer()
    Dim rsArchive As New ADODB.Recordset
    Dim Qry As String, QrySrch As String, Desc2() As String
    Dim i As Integer
    
    Select Case cboShow.ListIndex
        Case 0  'All
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
            QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE " & _
                    "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
        Case 1  'Applic Model
            '***Access
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.APPLICMODEL)='' Or (SPAREPARTS.APPLICMODEL) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.APPLICMODEL)='' Or (SPAREPARTS.APPLICMODEL) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE SPAREPARTS.APPLICMODEL='' Or (SPAREPARTS.APPLICMODEL) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE SPAREPARTS.APPLICMODEL='' Or (SPAREPARTS.APPLICMODEL) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
        
        Case 2  'Location
            '***Access
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.LOCATION)='' Or (SPAREPARTS.LOCATION) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.LOCATION)='' Or (SPAREPARTS.LOCATION) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE SPAREPARTS.LOCATION='' Or (SPAREPARTS.LOCATION) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE SPAREPARTS.LOCATION='' Or (SPAREPARTS.LOCATION) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            
        Case 3  'Other Description
            '***Access
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.DESCRIPTION2)='' Or (SPAREPARTS.DESCRIPTION2) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.DESCRIPTION2)='' Or (SPAREPARTS.DESCRIPTION2) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE SPAREPARTS.DESCRIPTION2='' Or (SPAREPARTS.DESCRIPTION2) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE SPAREPARTS.DESCRIPTION2='' Or (SPAREPARTS.DESCRIPTION2) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            
        Case 4  'Type
            '***Access
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.REMARKS)='' Or (SPAREPARTS.REMARKS) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.REMARKS)='' Or (SPAREPARTS.REMARKS) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE SPAREPARTS.REMARKS='' Or (SPAREPARTS.REMARKS) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE SPAREPARTS.REMARKS='' Or (SPAREPARTS.REMARKS) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            
        Case 5  'Superseding PN
            '***Access
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.SUPERSEDINGPN)='' Or (SPAREPARTS.SUPERSEDINGPN) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE (((SPAREPARTS.SUPERSEDINGPN)='' Or (SPAREPARTS.SUPERSEDINGPN) Is Null)) ORDER BY SPAREPARTS.PARTNUMBER;"
            'mySQL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.REMARKS, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.QTYONHAND, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.SPAREPARTID " & _
                  "From SPAREPARTS WHERE SPAREPARTS.SUPERSEDINGPN='' Or (SPAREPARTS.SUPERSEDINGPN) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            QrySrch = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.DESCRIPTION2, SPAREPARTS.LOCATION, SPAREPARTS.APPLICMODEL, SPAREPARTS.SUPERSEDINGPN, SPAREPARTS.REMARKS, SPAREPARTS.REORDERPOINT, SPAREPARTS.MAXALLOW, SPAREPARTS.QTYONHAND, SPAREPARTS.SUPPLIERCODE  " & _
                  "From SPAREPARTS WHERE SPAREPARTS.SUPERSEDINGPN='' Or (SPAREPARTS.SUPERSEDINGPN) Is Null ORDER BY SPAREPARTS.PARTNUMBER;"
            
    
    End Select
    If rsParts.State = adStateOpen Then rsParts.Close
    rsParts.Open QrySrch, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSrch.DataSource = rsParts
    With flxSrch
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = txtPartNo.Width
        For i = 3 To 11
            .ColWidth(i) = 0
        Next i
    End With
    
    If rsSP.State = adStateOpen Then rsSP.Close
    rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSP.RecordCount > 0 Then
        Set flxSP.DataSource = rsSP
    Else
        flxSP.Rows = 1
        flxSP.Rows = 2
    End If
    'Display records in Archive
    Qry = "SELECT ARCHIVESP.PARTNUMBER, ARCHIVESP.DESCRIPTION, ARCHIVESP.QTYONHAND, ARCHIVESP.DESCRIPTION2, ARCHIVESP.SPAREPARTID FROM ARCHIVESP ORDER BY PARTNUMBER;"
    With rsArchive
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxArchive.Rows = 1
            .MoveFirst
            Do While Not .EOF
                If Not IsNull(!Description2) Then
                    Desc2 = Split(!Description2, "|")
                    flxArchive.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & Desc2(0) & vbTab & Desc2(1) & vbTab & Desc2(2) & vbTab & Desc2(3) & vbTab & !SparepartID
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Timer1.Enabled = False
End Sub

Private Sub txtApplic_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub


Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDesc2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtLoc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtMax_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPartArc_Change()
    Dim Qry As String
    
    With FlxSrchArc
        If PrtSrch Or rsPrtsArc.RecordCount = 0 Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartArc.Height + txtPartArc.Top
        .Left = txtPartArc.Left
        .Width = txtPartArc.Width
        .Visible = True
        .ZOrder
        If txtPartArc.Text = "" Then .Visible = False: Exit Sub
        rsPrtsArc.Find "[PARTNUMBER] LIKE '" + txtPartArc.Text + "%'", , adSearchForward, 1
        If Not rsPrtsArc.EOF Then
            If Not PrtBln Then
                txtPartArc.Tag = Len(txtPartArc.Text)
                PrtBln = True
            End If
            If Not FromFlx Then     'top row only if user type @ txtPartArc
                .TopRow = rsPrtsArc.AbsolutePosition - 1
                FromFlx = False
            End If
            .Row = rsPrtsArc.AbsolutePosition - 1
            .Col = 0: .ColSel = 5
            txtPartArc.Text = rsPrtsArc!PartNumber
            flxArchive.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3)
            flxArchive.TextMatrix(ActRow, 3) = .TextMatrix(.Row, 4)
            flxArchive.TextMatrix(ActRow, 4) = Format(Date, "MM/DD/YYYY")
            flxArchive.TextMatrix(ActRow, 5) = .TextMatrix(.Row, 5)
            flxArchive.TextMatrix(ActRow, 8) = .TextMatrix(.Row, 1) 'spid
            txtPartArc.SelStart = Val(txtPartArc.Tag)
            txtPartArc.SelLength = Len(txtPartArc.Text)
            PrtBln = False
        Else
            flxArchive.TextMatrix(ActRow, 2) = ""
            flxArchive.TextMatrix(ActRow, 3) = ""
            flxArchive.TextMatrix(ActRow, 4) = ""
            flxArchive.TextMatrix(ActRow, 8) = ""
            flxArchive.TextMatrix(ActRow, 5) = ""
            .Visible = False
        End If
    End With
End Sub

Private Sub txtPartArc_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            txtPartArc.Tag = "skiplostfocus"
            FlxSrchArc.Visible = True
            FlxSrchArc.SetFocus
        Case vbKeyEscape
            txtPartArc.Text = Empty
            Tmp = Split(txtPartArc.Tag, "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxArchive.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
            txtPartArc.Tag = "skiplostfocus"
            flxArchive.SetFocus
            flxArchive.Col = 4
            txtPartArc.Visible = False
            FlxSrchArc.Visible = False
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
    End Select
End Sub

Private Sub txtPartArc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxArchive.SetFocus
        'Call txtPartNo_LostFocus
        'If all fields are filled then addrow focus on next entry
        If flxArchive.TextMatrix(ActRow, 8) <> "" And flxArchive.TextMatrix(ActRow, 2) <> "" Then
            If ActRow = flxArchive.Rows - 1 Then flxArchive.AddItem ""
            flxArchive.SetFocus
            flxArchive.Row = flxArchive.Rows - 1
            flxArchive.Col = 1
        Else
            flxArchive.SetFocus
            flxArchive.Col = 2
        End If
        FlxSrchArc.Visible = False
    End If

End Sub

Private Sub txtPartArc_LostFocus()
    Dim Qry As String, UDay As String
    Dim i As Integer
    
    If txtPartArc.Tag = "skiplostfocus" Then
        txtPartArc.Tag = ""
        Exit Sub
    End If
    If flxArchive.TextMatrix(ActRow, 8) = "" Or txtPartArc.Text = "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartArc.Visible = False
        FlxSrchArc.Visible = False
        For i = 1 To 8
            flxArchive.TextMatrix(ActRow, i) = ""
        Next i
        Exit Sub
    Else
        '***Access
        'Qry = "SELECT SPITR.TRANSDATE From SPITR " & _
              "WHERE (((SPITR.SPAREPARTID)=" & flxArchive.TextMatrix(ActRow, 8) & ") AND ((SPITR.TRANSTYPE)<>'CANCELLED' And (SPITR.TRANSTYPE)<>'ORDER TEMP')) ORDER BY SPITR.TRANSDATE DESC;"
        'mySQL
        Qry = "SELECT SPITR.TRANSDATE From SPITR " & _
              "WHERE (((SPITR.SPAREPARTID)=" & flxArchive.TextMatrix(ActRow, 8) & ") AND ((SPITR.TRANSTYPE)<>'CANCELLED' And (SPITR.TRANSTYPE)<>'ORDER TEMP')) ORDER BY SPITR.TRANSDATE DESC;"
        
        UDay = GetFieldValue(Qry)
        If IsDate(UDay) Then
            flxArchive.TextMatrix(ActRow, 6) = UDay
            flxArchive.TextMatrix(ActRow, 7) = DateDiff("D", Date, DateValue(UDay))
        Else
            flxArchive.TextMatrix(ActRow, 6) = "Unused"
            flxArchive.TextMatrix(ActRow, 7) = "Unknown"
        End If
        If flxArchive.TextMatrix(ActRow, 5) <> "" Then
            '***Access
            'Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & flxArchive.TextMatrix(ActRow, 5) & "'));"
            'mySQL
            Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & flxArchive.TextMatrix(ActRow, 5) & "';"
            flxArchive.TextMatrix(ActRow, 5) = GetFieldValue(Qry)
        End If
    End If
    flxArchive.TextMatrix(ActRow, 1) = txtPartArc.Text
    txtPartArc.Visible = False
    FlxSrchArc.Visible = False
    
End Sub
Private Sub txtPartNo_Change()
    Dim Qry As String
    
    With flxSrch
        If PrtSrch Or rsParts.RecordCount = 0 Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = Frame1.Top + txtPartNo.Height + txtPartNo.Top
        .Left = Frame1.Left + txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        .ZOrder
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
            .Col = 0: .ColSel = 5
            txtPartNo.Text = rsParts!PartNumber
            txtSuper.Text = .TextMatrix(.Row, 7) 'Superseding
            cboType.Text = .TextMatrix(.Row, 8) 'Remarks/Type
            txtReorder.Text = .TextMatrix(.Row, 9)  'Reorder
            txtMax.Text = .TextMatrix(.Row, 10) 'Max
            txtQty.Text = .TextMatrix(.Row, 11) 'Qty On hand
            txtDesc.Text = .TextMatrix(.Row, 3) 'Desc
            txtDesc2.Text = .TextMatrix(.Row, 4) 'Desc2
            txtLoc.Text = .TextMatrix(.Row, 5)  'Location
            txtApplic.Text = .TextMatrix(.Row, 6)   'ApplicModel
            txtSPID.Text = .TextMatrix(.Row, 1)
            txtSuppCode.Text = .TextMatrix(.Row, 12) 'Supplier Code
            If txtSuppCode.Text <> "" Then
                '***Access
                'Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE (((SUPPLIER.SUPPLIERCODE)='" & txtSuppCode.Text & "'));"
                'mySQL
                Qry = "SELECT SUPPLIER.SUPPLIERNAME From SUPPLIER WHERE SUPPLIER.SUPPLIERCODE='" & txtSuppCode.Text & "';"
                cboSupplier.Text = GetFieldValue(Qry)
            End If
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            txtSuper.Text = ""
            cboType.Text = ""
            txtReorder.Text = ""
            txtMax.Text = ""
            txtQty.Text = ""
            txtDesc.Text = ""
            txtDesc2.Text = ""
            txtLoc.Text = ""
            txtApplic.Text = ""
            txtSPID.Text = ""
            .Visible = False
        End If
    End With

End Sub

Private Sub txtPartNo_GotFocus()
    If txtPartNo.Text <> "" Then
        txtPartNo.SelStart = 0
        txtPartNo.SelLength = Len(txtPartNo.Text)
    End If
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            flxSrch.Tag = "skiplostfocus"
            flxSrch.Visible = True
            flxSrch.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            flxSrch.Visible = False
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
        Case Else
            FromFlx = False
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        flxSrch.Visible = False
        txtSuper.SetFocus
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    Dim i As Integer
    
    If flxSrch.Tag = "skiplostfocus" Then
        flxSrch.Tag = ""
        Exit Sub
    End If
    flxSrch.Visible = False
    If txtSPID.Text <> "" Then
        For i = 1 To flxSP.Rows - 1
            If flxSP.TextMatrix(i, 11) = txtSPID.Text Then
                ActRow = i
                flxSP.TopRow = i
                'flxSP.Row = i
                'flxSP.Col = 11: flxSP.Row = i: flxSP.ColSel = 1: flxSP.RowSel = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtReorder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtSuper_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtWithIN_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    
End Sub

Private Sub txtWithIN_LostFocus()
    If Val(txtWithIN.Text) > 0 And Val(txtWithIN.Text) <= 1000 Then updWin.Value = Val(txtWithIN.Text)
End Sub

Private Sub updMDY_DownClick()
    Select Case updMDY
        Case 1
            lblMDY.Caption = IIf(Val(txtWithIN.Text) > 1, "Months", "Month")
            updMDY.Value = 4
        Case 2
            lblMDY.Caption = IIf(Val(txtWithIN.Text) > 1, "Days", "Day")
        Case 3
            lblMDY.Caption = IIf(Val(txtWithIN.Text) > 1, "Years", "Year")
    End Select
End Sub

Private Sub updMDY_UpClick()
    Select Case updMDY
        Case 1
            lblMDY.Caption = IIf(Val(txtWithIN.Text) > 1, "Months", "Month")
        Case 2
            lblMDY.Caption = IIf(Val(txtWithIN.Text) > 1, "Days", "Day")
        Case 3
            lblMDY.Caption = IIf(Val(txtWithIN.Text) > 1, "Years", "Year")
            updMDY.Value = 0
    End Select
End Sub

Private Sub updWin_Change()
    txtWithIN.Text = updWin.Value
End Sub
