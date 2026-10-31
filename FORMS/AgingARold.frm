VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AgingAR 
   Caption         =   "Aging of Accounts Receivable"
   ClientHeight    =   9465
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   14445
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxInd 
      Height          =   2535
      Left            =   9000
      TabIndex        =   46
      Top             =   10080
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxComID 
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   9720
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab tabPeriod 
      Height          =   1575
      Left            =   4440
      TabIndex        =   20
      Top             =   9600
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Select Aging Period"
      TabPicture(0)   =   "AgingARold.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdOKI"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboNewPrd"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.ComboBox cboNewPrd 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3135
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   405
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   714
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
         Image           =   "AgingARold.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   1800
         TabIndex        =   23
         Top             =   1080
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   714
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
         Image           =   "AgingARold.frx":05B6
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdOKI 
         Height          =   405
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   714
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
         Image           =   "AgingARold.frx":0B50
         cBack           =   -2147483633
      End
   End
   Begin TabDlg.SSTab tabAging 
      Height          =   8415
      Left            =   120
      TabIndex        =   11
      Tag             =   "XY"
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Company"
      TabPicture(0)   =   "AgingARold.frx":10EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxAgingC"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Company Details"
      TabPicture(1)   =   "AgingARold.frx":1106
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxComDtls"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Individual"
      TabPicture(2)   =   "AgingARold.frx":1122
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxAgingI"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Print Preview"
      TabPicture(3)   =   "AgingARold.frx":113E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblPrnPrvw"
      Tab(3).Control(1)=   "CRViewer1"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   -74940
         TabIndex        =   38
         Tag             =   "XY"
         Top             =   320
         Width           =   14020
         Begin VB.ComboBox cboComDtls 
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
            Left            =   7920
            TabIndex        =   40
            Tag             =   "XYF"
            Top             =   172
            Width           =   2655
         End
         Begin VB.TextBox txtFindCD 
            Height          =   285
            Left            =   12000
            TabIndex        =   39
            Tag             =   "XYF"
            Top             =   195
            Width           =   1935
         End
         Begin lvButton.lvButtons_H cmdFindCD 
            Height          =   375
            Left            =   10680
            TabIndex        =   41
            Tag             =   "XY"
            Top             =   150
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
            Image           =   "AgingARold.frx":115A
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdPrnCD 
            Height          =   375
            Left            =   4560
            TabIndex        =   42
            Tag             =   "XY"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
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
            Image           =   "AgingARold.frx":16F4
            cBack           =   -2147483633
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Select Customer:"
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
            Left            =   6360
            TabIndex        =   45
            Top             =   210
            Width           =   1575
         End
         Begin VB.Label lblperiodCD 
            Caption         =   "For the Month of "
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
            TabIndex        =   43
            Top             =   210
            Width           =   4335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   -74940
         TabIndex        =   28
         Tag             =   "XY"
         Top             =   320
         Width           =   14020
         Begin VB.ComboBox cboIPeriod 
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
            Left            =   1680
            TabIndex        =   30
            Tag             =   "XYF"
            Top             =   172
            Width           =   3135
         End
         Begin VB.TextBox txtFindI 
            Height          =   285
            Left            =   12000
            TabIndex        =   29
            Tag             =   "XYF"
            Top             =   195
            Width           =   1935
         End
         Begin lvButton.lvButtons_H cmdFindI 
            Height          =   375
            Left            =   10680
            TabIndex        =   31
            Tag             =   "XY"
            Top             =   150
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
            Image           =   "AgingARold.frx":1C8E
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdINewAging 
            Height          =   375
            Left            =   4920
            TabIndex        =   32
            Tag             =   "XY"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "&New Aging"
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
            Image           =   "AgingARold.frx":2228
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdIReprocess 
            Height          =   375
            Left            =   6480
            TabIndex        =   33
            Tag             =   "XY"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "&Reprocess"
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
            Image           =   "AgingARold.frx":27C2
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdPrnI 
            Height          =   375
            Left            =   8040
            TabIndex        =   34
            Tag             =   "XY"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
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
            Image           =   "AgingARold.frx":2D5C
            cBack           =   -2147483633
         End
         Begin VB.Label Label2 
            Caption         =   "Aging Period:"
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
            TabIndex        =   35
            Top             =   210
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   60
         TabIndex        =   13
         Tag             =   "XY"
         Top             =   320
         Width           =   14020
         Begin VB.TextBox txtFind 
            Height          =   285
            Left            =   12000
            TabIndex        =   15
            Tag             =   "XYF"
            Top             =   195
            Width           =   1935
         End
         Begin VB.ComboBox cboCPeriod 
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
            Left            =   1680
            TabIndex        =   14
            Tag             =   "XYF"
            Top             =   172
            Width           =   3135
         End
         Begin lvButton.lvButtons_H cmdFindC 
            Height          =   375
            Left            =   10680
            TabIndex        =   16
            Tag             =   "XY"
            Top             =   150
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
            Image           =   "AgingARold.frx":32F6
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCNewAging 
            Height          =   375
            Left            =   4920
            TabIndex        =   17
            Tag             =   "XY"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "&New Aging"
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
            Image           =   "AgingARold.frx":3890
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCReprocess 
            Height          =   375
            Left            =   6480
            TabIndex        =   25
            Tag             =   "XY"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "&Reprocess"
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
            Image           =   "AgingARold.frx":3E2A
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdPrnC 
            Height          =   375
            Left            =   8040
            TabIndex        =   26
            Tag             =   "XY"
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
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
            Image           =   "AgingARold.frx":43C4
            cBack           =   -2147483633
         End
         Begin VB.Label Label1 
            Caption         =   "Aging Period:"
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
            TabIndex        =   18
            Top             =   210
            Width           =   1575
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAgingC 
         Height          =   7335
         Left            =   60
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   960
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   12938
         _Version        =   393216
         RowHeightMin    =   285
         BackColorBkg    =   16777215
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   7695
         Left            =   -74940
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   600
         Width           =   14055
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
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAgingI 
         Height          =   7335
         Left            =   -74940
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   960
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   12938
         _Version        =   393216
         RowHeightMin    =   285
         BackColorBkg    =   16777215
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxComDtls 
         Height          =   7335
         Left            =   -74940
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   960
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   12938
         _Version        =   393216
         RowHeightMin    =   285
         BackColorBkg    =   16777215
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblPrnPrvw 
         Caption         =   "Print Preview"
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   360
         Width           =   8295
      End
   End
   Begin VB.TextBox txtDueDate 
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   120
      TabIndex        =   7
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
      Interval        =   2000
      Left            =   6000
      Top             =   0
   End
   Begin VB.TextBox txtYr 
      Height          =   285
      Left            =   7800
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   6960
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   4
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
      Image           =   "AgingARold.frx":495E
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
      TabIndex        =   8
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
      m_MDI_Width     =   14565
      m_MDI_Height    =   9975
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHLY BILLING"
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
      Caption         =   "MONTHLY BILLING"
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
   Begin VB.Menu mnuIND 
      Caption         =   "Individual Aging"
      Visible         =   0   'False
      Begin VB.Menu mnuReProInd 
         Caption         =   "Reprocess Billing"
      End
   End
End
Attribute VB_Name = "AgingAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCAging As New ADODB.Recordset
Dim rsIAging As New ADODB.Recordset
Dim rsCDAging As New ADODB.Recordset
Dim rsPrn As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer, ProInd As Boolean, Reprocess As Boolean


Private Sub cboCPeriod_Click()
    Dim rsComDtl As New ADODB.Recordset
    Dim MoYr() As String, Qry As String
    
    If cboCPeriod.ListIndex = -1 Or cboCPeriod.Text = "" Then Exit Sub
    MoYr = Split(cboCPeriod.Text, "-")
    txtYr.Text = Trim(MoYr(1))
    txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & "1, " & txtYr.Text))
    txtDueDate.Text = DateAdd("m", 1, DateSerial(txtYr.Text, txtMo.Text, 1))
    lblperiodCD.Caption = "For the Month of " & cboCPeriod.Text
            
    Qry = "SELECT AGINGAR.CUSTOMERID, AGINGAR.AGEMONTH, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON, AGINGAR.REFNO, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
          "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='COMPANY')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    txtSQL.Text = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON, AGINGAR.REFNO, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
                  "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='COMPANY')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
                  
    If rsCAging.State = adStateOpen Then rsCAging.Close
    rsCAging.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxAgingC.DataSource = rsCAging
    flxAgingC.Refresh
    'Display Company Details
    Qry = "SELECT AGINGAR.CUSTOMERID, AGINGAR.AGEMONTH, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, AGINGAR.REFNO, AGINGAR.REFDATE, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
          "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='COMPDETAILS')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    txtSQL.Text = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, AGINGAR.REFNO, AGINGAR.REFDATE, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
                  "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='COMPDETAILS')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    If rsCDAging.State = adStateOpen Then rsCDAging.Close
    rsCDAging.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxComDtls.DataSource = rsCDAging
    flxComDtls.Refresh

    Call RXResizer1_ResizeComplete
    
    Qry = "SELECT DISTINCT CUSTOMER.CUSTOMERNAME FROM CUSTOMER INNER JOIN AGINGAR ON CUSTOMER.CUSTOMERID = AGINGAR.CUSTOMERID " & _
          "WHERE (((AGINGAR.AGETYPE)='COMPDETAILS') AND ((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ")) ORDER BY CUSTOMER.CUSTOMERNAME;"
    cboComDtls.Clear
    With rsComDtl
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboComDtls.AddItem "ALL COMPANY"
            Do While Not .EOF
                cboComDtls.AddItem !CustomerName
                .MoveNext
            Loop
        End If
        .Close
    End With
End Sub

Private Sub cboIPeriod_Click()
    Dim MoYr() As String
    
    If cboIPeriod.ListIndex = -1 Or cboIPeriod.Text = "" Then Exit Sub
    MoYr = Split(cboIPeriod.Text, "-")
    txtYr.Text = Trim(MoYr(1))
    txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & "1, " & txtYr.Text))
    txtDueDate.Text = DateAdd("m", 1, DateSerial(txtYr.Text, txtMo.Text, 1))
    ProInd = True
    Timer1.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    cmdOK.Visible = True
    cmdOK.Visible = False
    tabPeriod.Visible = False
    tabPeriod.Top = Pbar.Top + Pbar.Height
    tabAging.Enabled = True
End Sub

Private Sub cmdCNewAging_Click()
    Dim PDate As Date, i As Integer, j As Integer
    
    'display tab period
    With tabPeriod
        .Top = (Me.ScaleHeight - .Height) \ 2
        .Left = (Me.ScaleWidth - .Width) \ 2
        .Visible = True
        .ZOrder
    End With
    tabAging.Enabled = False
    PDate = DateSerial(Year(Date), Month(Date), 1)
    Do While Format(PDate, "mm/dd/yyyy") <> "01/01/2008"
        cboNewPrd.AddItem MonthName(Month(PDate)) & " - " & Year(PDate)
        PDate = DateAdd("m", -1, PDate)
    Loop
    'identify next aging period
    If cboCPeriod.ListCount <> 0 Then
        For i = 0 To cboCPeriod.ListCount - 1
            For j = 0 To cboNewPrd.ListCount - 1
                If cboCPeriod.List(i) = cboNewPrd.List(j) Then
                    cboNewPrd.RemoveItem j
                    Exit For
                End If
            Next j
        Next i
    End If
        
End Sub

Private Sub cmdCReprocess_Click()
    Dim Ans As Integer
    
    If cboCPeriod.ListIndex = -1 Then
        MsgBox "Please select an aging period to process!", vbExclamation + vbOKOnly, "Reprocess"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to reprocess aging for " & cboCPeriod.Text & "?", vbQuestion + vbYesNo, "Reprocess")
    If Ans = vbNo Then Exit Sub
    Reprocess = True
    cmdTrans(3).Enabled = False
    tabAging.Enabled = False
    Call cmdOK_Click
    Reprocess = False
    cmdTrans(3).Enabled = True
    tabAging.Enabled = True
End Sub

Private Sub cmdFindc_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxAgingC
        For i = fRow To .Rows - 1
            For j = fCol To 17
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 2: .Row = i
                    .ColSel = 17:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 17, i + 1, i)
                    fCol = IIf(j = 17, 2, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 2
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 2
End Sub


Private Sub cmdFindCD_Click()
    Dim i As Integer, j As Integer
        
    If txtFindCD.Text = "" Then Exit Sub
    With flxComDtls
        For i = fRow To .Rows - 1
            For j = fCol To 17
                If InStr(.TextMatrix(i, j), txtFindCD.Text) > 0 Then
                    .Col = 2: .Row = i
                    .ColSel = 17:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 17, i + 1, i)
                    fCol = IIf(j = 17, 2, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 2
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 2
End Sub

Private Sub cmdFindI_Click()
    Dim i As Integer, j As Integer
        
    If txtFindI.Text = "" Then Exit Sub
    With flxAgingI
        For i = fRow To .Rows - 1
            For j = fCol To 17
                If InStr(.TextMatrix(i, j), txtFindI.Text) > 0 Then
                    .Col = 2: .Row = i
                    .ColSel = 17:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 17, i + 1, i)
                    fCol = IIf(j = 17, 2, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 2
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 2

End Sub

Private Sub cmdINewAging_Click()
    Dim PDate As Date, i As Integer, j As Integer
    
    'display tab period
    With tabPeriod
        .Top = (Me.ScaleHeight - .Height) \ 2
        .Left = (Me.ScaleWidth - .Width) \ 2
        .Visible = True
        .ZOrder
    End With
    tabAging.Enabled = False
    PDate = DateSerial(Year(Date), Month(Date), 1)
    Do While Format(PDate, "mm/dd/yyyy") <> "01/01/2008"
        cboNewPrd.AddItem MonthName(Month(PDate)) & " - " & Year(PDate)
        PDate = DateAdd("m", -1, PDate)
    Loop
    'identify next aging period
    If cboIPeriod.ListCount <> 0 Then
        For i = 0 To cboIPeriod.ListCount - 1
            For j = 0 To cboNewPrd.ListCount - 1
                If cboIPeriod.List(i) = cboNewPrd.List(j) Then
                    cboNewPrd.RemoveItem j
                    Exit For
                End If
            Next j
        Next i
    End If
    cmdOK.Visible = False
    cmdOKI.Visible = True
End Sub

Private Sub cmdIReprocess_Click()
    Dim Ans As Integer
    
    If cboIPeriod.ListIndex = -1 Then
        MsgBox "Please select an aging period to process!", vbExclamation + vbOKOnly, "Reprocess"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to reprocess aging for " & cboIPeriod.Text & "?", vbQuestion + vbYesNo, "Reprocess")
    If Ans = vbNo Then Exit Sub
    Reprocess = True
    cmdTrans(3).Enabled = False
    tabAging.Enabled = False
    Call cmdOKI_Click
    Reprocess = False
    cmdTrans(3).Enabled = True
    tabAging.Enabled = True
End Sub

Private Sub cmdOK_Click()
    Dim rsAging As New ADODB.Recordset
    Dim MoYr() As String, Qry As String
    Dim i As Integer, Ans As Integer, pPct As Single
    
    If Not Reprocess Then
        If cboNewPrd.Text = "" Or cboNewPrd.ListIndex = -1 Then
            MsgBox "Please select Aging Period from list!", vbExclamation + vbOKOnly, "Validation"
            Exit Sub
        End If
        
        Ans = MsgBox("Do you want to process aging for " & cboNewPrd.Text & "?", vbQuestion + vbYesNo, "Process Aging")
        If Ans = vbNo Then
            Call cmdCancel_Click
            Exit Sub
        End If
        cmdTrans(3).Enabled = False
        tabPeriod.TabCaption(0) = "Processing Aging Period"
        tabPeriod.Enabled = False
        
        MoYr = Split(cboNewPrd.Text, "-")
    Else    'reprocess mode
        MoYr = Split(cboCPeriod.Text, "-")
    End If
    txtYr.Text = Trim(MoYr(1))
    txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & " 1, " & txtYr.Text))
    txtDueDate.Text = DateSerial(txtYr.Text, txtMo.Text, 1)
    
    Call ProcessCompanyAging
    DoEvents
    tabAging.Tab = 1
    DoEvents
    Call ProcessComDtlsAging
    Ans = MsgBox("Finish processing Aging for " & IIf(Reprocess, cboCPeriod.Text, cboNewPrd.Text) & "!" & vbCrLf & vbCrLf & "Do you want to Save Aging?", vbQuestion + vbYesNo, "Save")
    If Ans = vbNo Then
        cmdTrans(3).Enabled = True
        tabPeriod.TabCaption(0) = "Select Aging Period"
        tabPeriod.Enabled = True
        Call cmdCancel_Click
        Exit Sub
    End If
    If Reprocess Then
        'delete existing records
        LYRcon.Execute "DELETE AGINGAR.* From AgingAR WHERE (((AGINGAR.AGETYPE)='COMPANY') AND ((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & "));"
        LYRcon.Execute "DELETE AGINGAR.* From AgingAR WHERE (((AGINGAR.AGETYPE)='COMPDETAILS') AND ((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & "));"
    End If
    'Save Aging
    With flxAgingC
        Pbar.Value = 0
        Pbar.Max = (.Rows - 1) + (flxComDtls.Rows - 1)
        Pbar.Visible = True
        lblTotal.Caption = ""
        rsAging.Open "AgingAR", LYRcon, adOpenKeyset, adLockOptimistic
        For i = 1 To .Rows - 1
            rsAging.AddNew
            rsAging!CUSTOMERID = .TextMatrix(i, 1)
            rsAging!AGETYPE = "COMPANY"
            rsAging!RefNo = .TextMatrix(i, 6)
            rsAging!AgeMonth = txtMo.Text
            rsAging!AGEYEAR = txtYr.Text
            rsAging!DueNxtMo = Val(Replace(.TextMatrix(i, 7), ",", ""))
            rsAging!Current = Val(Replace(.TextMatrix(i, 8), ",", ""))
            rsAging!Due1to30 = Val(Replace(.TextMatrix(i, 9), ",", ""))
            rsAging!Due31to60 = Val(Replace(.TextMatrix(i, 10), ",", ""))
            rsAging!Due61to90 = Val(Replace(.TextMatrix(i, 11), ",", ""))
            rsAging!Due91UP = Val(Replace(.TextMatrix(i, 12), ",", ""))
            rsAging!TtlOverDue = Val(Replace(.TextMatrix(i, 13), ",", ""))
            rsAging!CollNxtMo = Val(Replace(.TextMatrix(i, 14), ",", ""))
            rsAging!DateDue = Val(.TextMatrix(i, 15))
            rsAging!DatePd = Val(.TextMatrix(i, 16))
            rsAging!AmtPd = Val(Replace(.TextMatrix(i, 17), ",", ""))
            rsAging.Update
            Pbar.Value = Pbar.Value + 1
            DoEvents
            pPct = (Pbar.Value / Pbar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
            DoEvents
        Next i
    End With
    With flxComDtls
        For i = 1 To .Rows - 1
            rsAging.AddNew
            rsAging!CUSTOMERID = .TextMatrix(i, 1)
            rsAging!AGETYPE = "COMPDETAILS"
            rsAging!RefNo = .TextMatrix(i, 5)
            rsAging!RefDate = .TextMatrix(i, 6)
            rsAging!AgeMonth = txtMo.Text
            rsAging!AGEYEAR = txtYr.Text
            rsAging!DueNxtMo = Val(Replace(.TextMatrix(i, 7), ",", ""))
            rsAging!Current = Val(Replace(.TextMatrix(i, 8), ",", ""))
            rsAging!Due1to30 = Val(Replace(.TextMatrix(i, 9), ",", ""))
            rsAging!Due31to60 = Val(Replace(.TextMatrix(i, 10), ",", ""))
            rsAging!Due61to90 = Val(Replace(.TextMatrix(i, 11), ",", ""))
            rsAging!Due91UP = Val(Replace(.TextMatrix(i, 12), ",", ""))
            rsAging!TtlOverDue = Val(Replace(.TextMatrix(i, 13), ",", ""))
            rsAging!CollNxtMo = Val(Replace(.TextMatrix(i, 14), ",", ""))
            rsAging!DateDue = Val(.TextMatrix(i, 15))
            rsAging!DatePd = Val(.TextMatrix(i, 16))
            rsAging!AmtPd = Val(Replace(.TextMatrix(i, 17), ",", ""))
            rsAging.Update
            Pbar.Value = Pbar.Value + 1
            DoEvents
            pPct = (Pbar.Value / Pbar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
            DoEvents
        Next i
    End With
    rsAging.Close
    If Not Reprocess Then
        cboCPeriod.AddItem cboNewPrd.Text, 0
        cboCPeriod.ListIndex = 0
    End If
    cmdTrans(3).Enabled = True
    tabPeriod.TabCaption(0) = "Select Aging Period"
    tabPeriod.Enabled = True
    Call cmdCancel_Click
    Call RXResizer1_ResizeComplete
    MsgBox "Aging for " & cboCPeriod.Text & " successfully saved!", vbInformation + vbOKOnly, "Save"
    Pbar.Visible = False
    lblTotal.Caption = ""
End Sub

Private Sub ProcessCompanyAging()
    Dim rsCID As New ADODB.Recordset
    Dim rsCust As New ADODB.Recordset
    Dim DueNxtM As Currency, Current As Currency, iRow As Integer, jRow As Integer, CoID As Integer, dPD As Integer
    Dim Due30 As Currency, Due60 As Currency, Due90 As Currency, Over90 As Currency, TtlOverDue As Currency, CNxtM As Currency, AmtPd As Currency
    Dim AgeStart As Date, AgeEnd As Date, dStart As Date, i As Integer, pPct As Single
    Dim Qry As String, Tmp As String
    
    AgeStart = txtDueDate.Text
    AgeEnd = DateAdd("d", -1, DateAdd("m", 1, AgeStart))
    Qry = "SELECT CUSTOMER.CUSTOMERID FROM CUSTOMER WHERE (((CUSTOMER.CTYPE)='COMPANY')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    rsCID.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxComID.DataSource = rsCID
    flxComID.Refresh
    'Start process
    If rsCAging.State = adStateOpen Then rsCAging.Close
    flxAgingC.Rows = 1
    Pbar.Value = 0
    Pbar.Max = flxComID.Rows - 1
    lblTotal.Caption = ""
    Pbar.Visible = True
    For i = 1 To flxComID.Rows - 1
        CoID = flxComID.TextMatrix(i, 1)
        Qry = "SELECT Count(ONACCTERMS.CUSTOMERID) AS CountOfCUSTOMERID From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
        ' skip customer without open account
        If Val(GetFieldValue(Qry)) > 0 Then
            Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON " & _
                  "From CUSTOMER WHERE (((CUSTOMER.CUSTOMERID)=" & CoID & "));"
            With rsCust
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    flxAgingC.AddItem vbTab & !CUSTOMERID & vbTab & flxAgingC.Rows & vbTab & !CustomerName & vbTab & !Address & vbTab & !ContactPerson & vbTab & "INVOICE"
                    iRow = flxAgingC.Rows - 1
                    flxAgingC.AddItem vbTab & !CUSTOMERID & vbTab & flxAgingC.Rows & vbTab & !CustomerName & vbTab & !Address & vbTab & !ContactPerson & vbTab & "JOB ORDER"
                    jRow = flxAgingC.Rows - 1
                End If
                .Close
            End With
            'compute details for invoice
            Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>=#" & AgeStart & "# And (ONACCTERMS.TRANSDATE)<=#" & AgeEnd & "#) AND ((ONACCTERMS.JOBORDERID)=0));"
            DueNxtM = Val(GetFieldValue(Qry))
            Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>#" & AgeEnd & "#) AND ((ONACCTERMS.JOBORDERID)=0));"
            Current = Val(GetFieldValue(Qry))
            dStart = DateAdd("m", -1, AgeStart)
            Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & DateAdd("d", -1, DateAdd("m", 1, dStart)) & "#) AND ((ONACCTERMS.JOBORDERID)=0));"
            Due30 = Val(GetFieldValue(Qry))
            dStart = DateAdd("m", -2, AgeStart)
            Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & DateAdd("d", -1, DateAdd("m", 1, dStart)) & "#) AND ((ONACCTERMS.JOBORDERID)=0));"
            Due60 = Val(GetFieldValue(Qry))
            dStart = DateAdd("m", -3, AgeStart)
            Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & DateAdd("d", -1, DateAdd("m", 1, dStart)) & "#) AND ((ONACCTERMS.JOBORDERID)=0));"
            Due90 = Val(GetFieldValue(Qry))
            Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)<#" & dStart & "#) AND ((ONACCTERMS.JOBORDERID)=0));"
            Over90 = Val(GetFieldValue(Qry))
            TtlOverDue = Due30 + Due60 + Due90 + Over90
            CNxtM = DueNxtM + TtlOverDue
            Qry = "SELECT LEDGER.REFDATE From Ledger WHERE (((LEDGER.REFDATE)>=#" & AgeStart & "# And (LEDGER.REFDATE)<=#" & AgeEnd & "#) AND ((LEDGER.CUSTOMERID)=" & CoID & ") AND ((LEDGER.REFTYPE)='Payment') AND ((LEDGER.JOBID)=0)) ORDER BY LEDGER.REFDATE DESC;"
            Tmp = GetFieldValue(Qry)
            If Tmp <> "" Then
                dPD = Day(DateValue(Tmp))
            Else
                dPD = 0
            End If
            Qry = "SELECT Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger " & _
                  "WHERE (((LEDGER.REFDATE)>=#" & AgeStart & "# And (LEDGER.REFDATE)<=#" & AgeEnd & "#) AND ((LEDGER.CUSTOMERID)=" & CoID & ") AND ((LEDGER.REFTYPE)='Payment') AND ((LEDGER.JOBID)=0));"
            AmtPd = Val(GetFieldValue(Qry))
            'Display in flx
            With flxAgingC
                .TextMatrix(iRow, 7) = Format(DueNxtM, "#,##0.00")
                .TextMatrix(iRow, 8) = Format(Current, "#,##0.00")
                .TextMatrix(iRow, 9) = Format(Due30, "#,##0.00")
                .TextMatrix(iRow, 10) = Format(Due60, "#,##0.00")
                .TextMatrix(iRow, 11) = Format(Due90, "#,##0.00")
                .TextMatrix(iRow, 12) = Format(Over90, "#,##0.00")
                .TextMatrix(iRow, 13) = Format(TtlOverDue, "#,##0.00")
                .TextMatrix(iRow, 14) = Format(CNxtM, "#,##0.00")
                .TextMatrix(iRow, 15) = 15
                .TextMatrix(iRow, 16) = dPD
                .TextMatrix(iRow, 17) = Format(AmtPd, "#,##0.00")
            End With
            Qry = "SELECT Count(ONACCTERMS.CUSTOMERID) AS CountOfCUSTOMERID From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.JOBORDERID)>0));"
            If Val(GetFieldValue(Qry)) > 0 Then
                'compute details for job order
                Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>=#" & AgeStart & "# And (ONACCTERMS.TRANSDATE)<=#" & AgeEnd & "#) AND ((ONACCTERMS.JOBORDERID)>0));"
                DueNxtM = Val(GetFieldValue(Qry))
                Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>#" & AgeEnd & "#) AND ((ONACCTERMS.JOBORDERID)>0));"
                Current = Val(GetFieldValue(Qry))
                dStart = DateAdd("m", -1, AgeStart)
                Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & DateAdd("d", -1, DateAdd("m", 1, dStart)) & "#) AND ((ONACCTERMS.JOBORDERID)>0));"
                Due30 = Val(GetFieldValue(Qry))
                dStart = DateAdd("m", -2, AgeStart)
                Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & DateAdd("d", -1, DateAdd("m", 1, dStart)) & "#) AND ((ONACCTERMS.JOBORDERID)>0));"
                Due60 = Val(GetFieldValue(Qry))
                dStart = DateAdd("m", -3, AgeStart)
                Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & DateAdd("d", -1, DateAdd("m", 1, dStart)) & "#) AND ((ONACCTERMS.JOBORDERID)>0));"
                Due90 = Val(GetFieldValue(Qry))
                Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
                      "WHERE (((ONACCTERMS.CUSTOMERID)=" & CoID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.TRANSDATE)<#" & dStart & "#) AND ((ONACCTERMS.JOBORDERID)>0));"
                Over90 = Val(GetFieldValue(Qry))
                TtlOverDue = Due30 + Due60 + Due90 + Over90
                CNxtM = DueNxtM + TtlOverDue
                Qry = "SELECT LEDGER.REFDATE From Ledger WHERE (((LEDGER.REFDATE)>=#" & AgeStart & "# And (LEDGER.REFDATE)<=#" & AgeEnd & "#) AND ((LEDGER.CUSTOMERID)=" & CoID & ") AND ((LEDGER.REFTYPE)='Payment') AND ((LEDGER.JOBID)>0)) ORDER BY LEDGER.REFDATE DESC;"
                Tmp = GetFieldValue(Qry)
                If Tmp <> "" Then
                    dPD = Day(DateValue(Tmp))
                Else
                    dPD = 0
                End If
                Qry = "SELECT Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger " & _
                      "WHERE (((LEDGER.REFDATE)>=#" & AgeStart & "# And (LEDGER.REFDATE)<=#" & AgeEnd & "#) AND ((LEDGER.CUSTOMERID)=" & CoID & ") AND ((LEDGER.REFTYPE)='Payment') AND ((LEDGER.JOBID)>0));"
                AmtPd = Val(GetFieldValue(Qry))
                'Display in flx
                With flxAgingC
                    .TextMatrix(jRow, 7) = Format(DueNxtM, "#,##0.00")
                    .TextMatrix(jRow, 8) = Format(Current, "#,##0.00")
                    .TextMatrix(jRow, 9) = Format(Due30, "#,##0.00")
                    .TextMatrix(jRow, 10) = Format(Due60, "#,##0.00")
                    .TextMatrix(jRow, 11) = Format(Due90, "#,##0.00")
                    .TextMatrix(jRow, 12) = Format(Over90, "#,##0.00")
                    .TextMatrix(jRow, 13) = Format(TtlOverDue, "#,##0.00")
                    .TextMatrix(jRow, 14) = Format(CNxtM, "#,##0.00")
                    .TextMatrix(jRow, 15) = 15
                    .TextMatrix(jRow, 16) = dPD
                    .TextMatrix(jRow, 17) = Format(AmtPd, "#,##0.00")
                End With
            Else
                flxAgingC.RemoveItem jRow
            End If
        End If
        Pbar.Value = Pbar.Value + 1
        DoEvents
        pPct = (Pbar.Value / Pbar.Max)
        lblTotal.Caption = Format$(pPct, "##0.0%")
        DoEvents
    Next i
    'Pbar.Visible = False
    lblTotal.Caption = ""
    rsCID.Close
End Sub

Private Sub cmdOKI_Click()
    Dim rsAging As New ADODB.Recordset
    Dim MoYr() As String, Qry As String
    Dim i As Integer, Ans As Integer, pPct As Single
    
    If Not Reprocess Then
        If cboNewPrd.Text = "" Or cboNewPrd.ListIndex = -1 Then
            MsgBox "Please select Aging Period from list!", vbExclamation + vbOKOnly, "Validation"
            Exit Sub
        End If
        
        Ans = MsgBox("Do you want to process aging for " & cboNewPrd.Text & "?", vbQuestion + vbYesNo, "Process Aging")
        If Ans = vbNo Then
            Call cmdCancel_Click
            Exit Sub
        End If
        cmdTrans(3).Enabled = False
        tabPeriod.TabCaption(0) = "Processing Aging Period"
        tabPeriod.Enabled = False
        
        MoYr = Split(cboNewPrd.Text, "-")
    Else
        MoYr = Split(cboIPeriod.Text, "-")
    End If
    txtYr.Text = Trim(MoYr(1))
    txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & " 1, " & txtYr.Text))
    txtDueDate.Text = DateSerial(txtYr.Text, txtMo.Text, 1)
    
    Call ProcessIndividualAging
    
    Ans = MsgBox("Finish processing Aging for " & IIf(Reprocess, cboIPeriod.Text, cboNewPrd) & "!" & vbCrLf & vbCrLf & "Do you want to Save Aging?", vbQuestion + vbYesNo, "Save")
    If Ans = vbNo Then
        cmdTrans(3).Enabled = True
        tabPeriod.TabCaption(0) = "Select Aging Period"
        tabPeriod.Enabled = True
        Call cmdCancel_Click
        Exit Sub
    End If
    'Save Aging
    If Reprocess Then LYRcon.Execute "DELETE AGINGAR.* From AgingAR WHERE (((AGINGAR.AGETYPE)='INDIVIDUAL') AND ((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & "));"
            
    With flxAgingI
        Pbar.Value = 0
        Pbar.Max = .Rows - 1
        Pbar.Visible = True
        lblTotal.Caption = ""
        rsAging.Open "AgingAR", LYRcon, adOpenKeyset, adLockOptimistic
        For i = 1 To .Rows - 1
            rsAging.AddNew
            rsAging!CUSTOMERID = .TextMatrix(i, 1)
            rsAging!AGETYPE = "INDIVIDUAL"
            rsAging!RefNo = .TextMatrix(i, 5)
            rsAging!RefDate = .TextMatrix(i, 6)
            rsAging!AgeMonth = txtMo.Text
            rsAging!AGEYEAR = txtYr.Text
            rsAging!Term = .TextMatrix(i, 7)
            rsAging!InvAge = Val(.TextMatrix(i, 8))
            rsAging!RefAmt = Val(Replace(.TextMatrix(i, 9), ",", ""))
            rsAging!DueNxtMo = Val(Replace(.TextMatrix(i, 10), ",", ""))
            rsAging!Current = Val(Replace(.TextMatrix(i, 11), ",", ""))
            rsAging!Due1to30 = Val(Replace(.TextMatrix(i, 12), ",", ""))
            rsAging!Due31to60 = Val(Replace(.TextMatrix(i, 13), ",", ""))
            rsAging!Due61to90 = Val(Replace(.TextMatrix(i, 14), ",", ""))
            rsAging!Due91UP = Val(Replace(.TextMatrix(i, 15), ",", ""))
            rsAging!TtlOverDue = Val(Replace(.TextMatrix(i, 16), ",", ""))
            rsAging!CollNxtMo = Val(Replace(.TextMatrix(i, 17), ",", ""))
            rsAging!BALANCE = Val(Replace(.TextMatrix(i, 18), ",", ""))
            rsAging!DatePd = Val(.TextMatrix(i, 19))
            rsAging!AmtPd = Val(Replace(.TextMatrix(i, 20), ",", ""))
            rsAging.Update
            Pbar.Value = Pbar.Value + 1
            DoEvents
            pPct = (Pbar.Value / Pbar.Max)
            lblTotal.Caption = Format$(pPct, "##0.0%")
            DoEvents
        Next i
        rsAging.Close
    End With
    If Not Reprocess Then
        cboIPeriod.AddItem cboNewPrd.Text, 0
        cboIPeriod.ListIndex = 0
    End If
    cmdTrans(3).Enabled = True
    tabPeriod.TabCaption(0) = "Select Aging Period"
    tabPeriod.Enabled = True
    Call cmdCancel_Click
    ProInd = True
    Timer1.Enabled = True
    MsgBox "Aging for " & cboIPeriod.Text & " successfully saved!", vbInformation + vbOKOnly, "Save"
    Pbar.Visible = False
    lblTotal.Caption = ""
End Sub

Private Sub ProcessIndividualAging()
    Dim rsCID As New ADODB.Recordset
    Dim rsCust As New ADODB.Recordset
    Dim jRow As Integer, CustID As Integer, dPD As Integer, MyNP As Integer, ActAmt As Currency, InvAge As Integer, Bal As Currency, MDue As Currency, j As Integer, MyDDif As Integer
    Dim Due30 As Currency, Due60 As Currency, Due90 As Currency, Over90 As Currency, TtlOverDue As Currency, Currnt As Currency, AmtPd As Currency, AmtPdOrig As Currency
    Dim AgeStart As Date, AgeEnd As Date, dTrans As Date, dStart As Date, i As Integer, pPct As Single, mTerm As Integer, DnxtM As Currency, TmpMDue As Currency, TmpDueDate As Date, tmpBal As Currency
    Dim Qry As String, Tmp As String, MyTerm As String
    
    AgeStart = txtDueDate.Text
    AgeEnd = DateAdd("d", -1, DateAdd("m", 1, AgeStart))
    Qry = "SELECT CUSTOMER.CUSTOMERID FROM CUSTOMER WHERE (((CUSTOMER.CTYPE)='INDIVIDUAL')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    If rsCID.State = adStateOpen Then rsCID.Close
    rsCID.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxComID.DataSource = rsCID
    flxComID.Refresh
    'Start process
    If rsIAging.State = adStateOpen Then rsIAging.Close
    flxAgingI.Rows = 1
    Pbar.Value = 0
    Pbar.Max = flxComID.Rows - 1
    lblTotal.Caption = ""
    Pbar.Visible = True
    For i = 1 To flxComID.Rows - 1
        CustID = flxComID.TextMatrix(i, 1)
        Qry = "SELECT Count(ONACCTERMS.CUSTOMERID) AS CountOfCUSTOMERID From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
        ' skip customer without open account
        If Val(GetFieldValue(Qry)) > 0 Then
            Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, ONACCTERMS.* FROM CUSTOMER INNER JOIN ONACCTERMS ON CUSTOMER.CUSTOMERID = ONACCTERMS.CUSTOMERID " & _
                  "WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.TRANSDATE)<=#" & AgeEnd & "#) AND ((ONACCTERMS.STATUS)='OPEN')) ORDER BY CUSTOMER.CUSTOMERNAME, ONACCTERMS.JOBORDERID, ONACCTERMS.TRANSDATE;"
            With rsCust
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        flxAgingI.AddItem vbTab & !CUSTOMERID & vbTab & flxAgingI.Rows & vbTab & !CustomerName & vbTab & !Address & vbTab & IIf(!JOBORDERID > 0, "JO# " & !JOBORDERNO, "Inv# " & !SPINVNO) & vbTab & Format(!TRANSDATE, "mm/dd/yyyy")
                        jRow = flxAgingI.Rows - 1
                        dTrans = !TRANSDATE
                        If Not IsNull(rsCust!Npays) Then
                            MyNP = rsCust!Npays
                            If Not IsNull(rsCust!PayScheme) Then
                                MyTerm = MyNP & "-" & Left$(rsCust!PayScheme, 1)
                            Else
                                MyTerm = MyNP
                            End If
                        Else
                            MyNP = 0
                            If Not IsNull(rsCust!PayScheme) Then
                                MyTerm = Left$(rsCust!PayScheme, 1)
                            Else
                                MyTerm = ""
                            End If
                        End If
                        InvAge = DateDiff("d", dTrans, DateSerial(Year(txtDueDate.Text), Month(txtDueDate.Text), Day(dTrans)))
                        ActAmt = IIf(IsNull(rsCust!AcctAmount), 0, rsCust!AcctAmount)
                        If !JOBORDERID > 0 Then
                            Tmp = "SELECT Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger WHERE (((LEDGER.SPINVNO)='" & !JOBORDERNO & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ") AND ((LEDGER.JOBID)>0));"
                        Else
                            Tmp = "SELECT Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger WHERE (((LEDGER.SPINVNO)='" & !SPINVNO & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ") AND ((LEDGER.JOBID) Is Null Or (LEDGER.JOBID)=0));"
                        End If
                        AmtPd = Val(GetFieldValue(Tmp))
                        AmtPdOrig = AmtPd
                        Bal = IIf(IsNull(!BALANCE), 0, !BALANCE)
                        If Not IsNull(!PayScheme) Then
                            If !PayScheme = "WEEKLY" And MyNP > 0 Then
                                MDue = IIf(MyNP >= 4, (ActAmt / MyNP) * 4, ActAmt)
                                mTerm = IIf(MyNP >= 4, MyNP \ 4, 1)
                                If mTerm > 1 And (MyNP / 4) > mTerm Then mTerm = mTerm + 1
                            ElseIf !PayScheme = "BI-MONTHLY" And MyNP > 0 Then
                                MDue = IIf(MyNP >= 2, (ActAmt / MyNP) * 2, ActAmt)
                                mTerm = IIf(MyNP >= 2, MyNP \ 2, 1)
                                If mTerm > 1 And (MyNP / 2) > mTerm Then mTerm = mTerm + 1
                            ElseIf !PayScheme = "MONTHLY" And MyNP > 0 Then
                                MDue = ActAmt / MyNP
                                mTerm = MyNP
                            End If
                        Else
                            Exit Do
                        End If
                        'initialize value of variables.
                        DnxtM = 0: Currnt = 0: Due30 = 0: Due60 = 0: Due90 = 0: Over90 = 0
                        dStart = DateSerial(Year(AgeStart), Month(AgeStart), Day(dTrans))
                        For j = 1 To mTerm + 1
                            TmpMDue = IIf(ActAmt > MDue, MDue, ActAmt)
                            ActAmt = ActAmt - TmpMDue
                            If j = 1 Then
                                TmpDueDate = DateAdd("m", 1, dTrans)
                            Else
                                TmpDueDate = DateAdd("m", 1, TmpDueDate)
                            End If
                            tmpBal = IIf(AmtPd > TmpMDue, 0, TmpMDue - AmtPd)
                            If AmtPd > 0 Then AmtPd = AmtPd - (TmpMDue - tmpBal)
                            MyDDif = DateDiff("m", TmpDueDate, dStart)
                            Select Case MyDDif
                                Case Is < 0
                                    Currnt = Currnt + tmpBal
                                Case 0
                                    DnxtM = DnxtM + tmpBal
                                Case 1
                                    Due30 = Due30 + tmpBal
                                Case 2
                                    Due60 = Due60 + tmpBal
                                Case 3
                                    Due90 = Due90 + tmpBal
                                Case Is > 3
                                    Over90 = Over90 + tmpBal
                            End Select
                        Next j
                        If !JOBORDERID > 0 Then
                            Qry = "SELECT Ledger.REFDATE From Ledger WHERE (((Ledger.REFDATE)>=#" & AgeStart & "# And (Ledger.REFDATE)<=#" & AgeEnd & "#) AND ((Ledger.CUSTOMERID)=" & CustID & ") AND ((Ledger.REFTYPE)='Payment') AND ((Ledger.JOBID)>0) AND ((Ledger.SPINVNO)='" & !JOBORDERNO & "')) ORDER BY Ledger.REFDATE DESC;"
                        Else
                            Qry = "SELECT Ledger.REFDATE From Ledger WHERE (((Ledger.REFDATE)>=#" & AgeStart & "# And (Ledger.REFDATE)<=#" & AgeEnd & "#) AND ((Ledger.CUSTOMERID)=" & CustID & ") AND ((Ledger.REFTYPE)='Payment') AND ((Ledger.JOBID)=0 Or (Ledger.JOBID) Is Null) AND ((Ledger.SPINVNO)='" & !SPINVNO & "')) ORDER BY Ledger.REFDATE DESC;"
                        End If
                        Tmp = GetFieldValue(Qry)
                        If Tmp <> "" Then
                            dPD = Day(DateValue(Tmp))
                        Else
                            dPD = 0
                        End If
                        flxAgingI.TextMatrix(jRow, 7) = MyTerm 'Term
                        flxAgingI.TextMatrix(jRow, 8) = InvAge  'Invoice Age
                        flxAgingI.TextMatrix(jRow, 9) = IIf(IsNull(rsCust!AcctAmount), "0.00", Format(rsCust!AcctAmount, "#,##0.00"))
                        flxAgingI.TextMatrix(jRow, 10) = Format(DnxtM, "#,##0.00")  'Due Next MonthIIf(IsNull(rsCust!AcctAmount), 0, rsCust!AcctAmount)
                        flxAgingI.TextMatrix(jRow, 11) = Format(Currnt, "#,##0.00")
                        flxAgingI.TextMatrix(jRow, 12) = Format(Due30, "#,##0.00")
                        flxAgingI.TextMatrix(jRow, 13) = Format(Due60, "#,##0.00")
                        flxAgingI.TextMatrix(jRow, 14) = Format(Due90, "#,##0.00")
                        flxAgingI.TextMatrix(jRow, 15) = Format(Over90, "#,##0.00")
                        flxAgingI.TextMatrix(jRow, 16) = Format(Due30 + Due60 + Due90 + Over90, "#,##0.00")
                        flxAgingI.TextMatrix(jRow, 17) = Format(Due30 + Due60 + Due90 + Over90 + DnxtM, "#,##0.00")
                        flxAgingI.TextMatrix(jRow, 18) = Format(Bal, "#,##0.00")
                        flxAgingI.TextMatrix(jRow, 19) = dPD
                        flxAgingI.TextMatrix(jRow, 20) = Format(AmtPdOrig, "#,##0.00")
                        .MoveNext
                    Loop
                End If
                .Close
            End With
        End If
        Pbar.Value = Pbar.Value + 1
        DoEvents
        pPct = (Pbar.Value / Pbar.Max)
        lblTotal.Caption = Format$(pPct, "##0.0%")
        DoEvents
    Next i
    Pbar.Visible = False
    lblTotal.Caption = ""
    rsCID.Close
End Sub

Private Sub ProcessComDtlsAging()
    Dim rsCID As New ADODB.Recordset
    Dim rsCust As New ADODB.Recordset
    Dim jRow As Integer, CustID As Integer, dPD As Integer, rCount As Integer
    Dim Due30 As Currency, Due60 As Currency, Due90 As Currency, Over90 As Currency, TtlOverDue As Currency, CNxtM As Currency, AmtPd As Currency
    Dim AgeStart As Date, AgeEnd As Date, dTrans As Date, dStart As Date, i As Integer, pPct As Single
    Dim Qry As String, Tmp As String
    
    AgeStart = txtDueDate.Text
    AgeEnd = DateAdd("d", -1, DateAdd("m", 1, AgeStart))
    Qry = "SELECT CUSTOMER.CUSTOMERID FROM CUSTOMER WHERE (((CUSTOMER.CTYPE)='COMPANY')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    If rsCID.State = adStateOpen Then rsCID.Close
    rsCID.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxComID.DataSource = rsCID
    flxComID.Refresh
    'Start process
    If rsCDAging.State = adStateOpen Then rsCDAging.Close
    'flxComDtls.Clear  'Clear FlxGrid here
    flxComDtls.Rows = 2
    'flxComDtls.FormatString = "   |  |^ No. |<    CUSTOMER     |<  ADDRESS   |<     REF. NO.  |<REF. DATE|>DUE NEXT MO. |>CURRENT|>1-30 OVERDUE|>31-60 OVERDUE |>61-90 OVERDUE|>91 & UP OVERDUE|>TOTAL OVERDUE|>COLLECTABLE|^DATE DUE|^DATE PAID|>AMOUNT PAID"
    Pbar.Value = 0
    Pbar.Visible = True
    Qry = "SELECT Count(CUSTOMER.CUSTOMERNAME) AS CountOfCUSTOMERNAME FROM CUSTOMER INNER JOIN ONACCTERMS ON CUSTOMER.CUSTOMERID = ONACCTERMS.CUSTOMERID " & _
          "WHERE (((ONACCTERMS.STATUS)='OPEN') AND ((CUSTOMER.CTYPE)='COMPANY'));"
    Pbar.Max = Val(GetFieldValue(Qry))
    lblTotal.Caption = ""
    Pbar.Visible = True
    rCount = 1
    For i = 1 To flxComID.Rows - 1
        CustID = flxComID.TextMatrix(i, 1)
        Qry = "SELECT Count(ONACCTERMS.CUSTOMERID) AS CountOfCUSTOMERID From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
        ' skip customer without open account
        If Val(GetFieldValue(Qry)) > 0 Then
            Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, ONACCTERMS.SPINVNO, ONACCTERMS.JOBORDERNO, ONACCTERMS.JOBORDERID, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.STATUS " & _
                  "FROM CUSTOMER INNER JOIN ONACCTERMS ON CUSTOMER.CUSTOMERID = ONACCTERMS.CUSTOMERID WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.STATUS)='OPEN')) ORDER BY ONACCTERMS.JOBORDERID, ONACCTERMS.TRANSDATE;"
            With rsCust
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        flxComDtls.AddItem vbTab & !CUSTOMERID & vbTab & rCount & vbTab & !CustomerName & vbTab & !Address & vbTab & IIf(!JOBORDERID > 0, "JO# " & !JOBORDERNO, "Inv# " & !SPINVNO) & vbTab & Format(!TRANSDATE, "mm/dd/yyyy")
                        rCount = rCount + 1
                        jRow = flxComDtls.Rows - 1
                        If IsNull(!TRANSDATE) Then
                            If IsNull(!SPINVNO) Then
                                Qry = "SELECT SRVCJOBORDER.DATEPROMISED From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & !JOBORDERNO & "'));"
                            Else
                                Qry = "SELECT SPINVOICE.INVDATE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & !SPINVNO & "'));"
                            End If
                            Tmp = GetFieldValue(Qry)
                            dTrans = IIf(IsDate(Tmp), DateValue(Tmp), "8/8/2008")
                        Else
                            dTrans = !TRANSDATE
                        End If
                        If dTrans >= AgeStart And dTrans <= AgeEnd Then
                            flxComDtls.TextMatrix(jRow, 7) = Format(!BALANCE, "#,##0.00") 'Due next month
                            flxComDtls.TextMatrix(jRow, 14) = Format(!BALANCE, "#,##0.00")  'Collectable
                        ElseIf dTrans > AgeEnd Then
                            flxComDtls.TextMatrix(jRow, 8) = Format(!BALANCE, "#,##0.00")  'current
                        Else
                            dStart = DateAdd("m", -1, AgeStart) 'Due 1 to 30
                            If dTrans >= dStart And dTrans <= DateAdd("d", -1, DateAdd("m", 1, dStart)) Then flxComDtls.TextMatrix(jRow, 9) = Format(!BALANCE, "#,##0.00"): flxComDtls.TextMatrix(jRow, 13) = Format(!BALANCE, "#,##0.00")
                            dStart = DateAdd("m", -2, AgeStart) 'Due 31 to 60
                            If dTrans >= dStart And dTrans <= DateAdd("d", -1, DateAdd("m", 1, dStart)) Then flxComDtls.TextMatrix(jRow, 10) = Format(!BALANCE, "#,##0.00"): flxComDtls.TextMatrix(jRow, 13) = Format(!BALANCE, "#,##0.00")
                            dStart = DateAdd("m", -3, AgeStart) 'Due 61 to 90
                            If dTrans >= dStart And dTrans <= DateAdd("d", -1, DateAdd("m", 1, dStart)) Then flxComDtls.TextMatrix(jRow, 11) = Format(!BALANCE, "#,##0.00"): flxComDtls.TextMatrix(jRow, 13) = Format(!BALANCE, "#,##0.00")
                            If dTrans < dStart Then flxComDtls.TextMatrix(jRow, 12) = Format(!BALANCE, "#,##0.00"): flxComDtls.TextMatrix(jRow, 13) = Format(!BALANCE, "#,##0.00")
                            flxComDtls.TextMatrix(jRow, 14) = Format(!BALANCE, "#,##0.00")
                            flxComDtls.TextMatrix(jRow, 15) = 15    'Day(dTrans)
                            
                        End If
                        If !JOBORDERID > 0 Then
                            Qry = "SELECT Ledger.REFDATE From Ledger WHERE (((Ledger.REFDATE)>=#" & AgeStart & "# And (Ledger.REFDATE)<=#" & AgeEnd & "#) AND ((Ledger.CUSTOMERID)=" & CustID & ") AND ((Ledger.REFTYPE)='Payment') AND ((Ledger.JOBID)>0) AND ((Ledger.SPINVNO)='" & !JOBORDERNO & "')) ORDER BY Ledger.REFDATE DESC;"
                        Else
                            Qry = "SELECT Ledger.REFDATE From Ledger WHERE (((Ledger.REFDATE)>=#" & AgeStart & "# And (Ledger.REFDATE)<=#" & AgeEnd & "#) AND ((Ledger.CUSTOMERID)=" & CustID & ") AND ((Ledger.REFTYPE)='Payment') AND ((Ledger.JOBID)=0 Or (Ledger.JOBID) Is Null) AND ((Ledger.SPINVNO)='" & !SPINVNO & "')) ORDER BY Ledger.REFDATE DESC;"
                        End If
                        Tmp = GetFieldValue(Qry)
                        If Tmp <> "" Then
                            dPD = Day(DateValue(Tmp))
                        Else
                            dPD = 0
                        End If
                        flxComDtls.TextMatrix(jRow, 16) = dPD
                        If !JOBORDERID > 0 Then
                            Qry = "SELECT Ledger.CREDIT From Ledger WHERE (((Ledger.REFDATE)>=#" & AgeStart & "# And (Ledger.REFDATE)<=#" & AgeEnd & "#) AND ((Ledger.CUSTOMERID)=" & CustID & ") AND ((Ledger.REFTYPE)='Payment') AND ((Ledger.JOBID)>0) AND ((Ledger.SPINVNO)='" & !JOBORDERNO & "')) ORDER BY Ledger.REFDATE DESC;"
                        Else
                            Qry = "SELECT Ledger.CREDIT From Ledger WHERE (((Ledger.REFDATE)>=#" & AgeStart & "# And (Ledger.REFDATE)<=#" & AgeEnd & "#) AND ((Ledger.CUSTOMERID)=" & CustID & ") AND ((Ledger.REFTYPE)='Payment') AND ((Ledger.JOBID)=0 Or (Ledger.JOBID) Is Null) AND ((Ledger.SPINVNO)='" & !SPINVNO & "')) ORDER BY Ledger.REFDATE DESC;"
                        End If
                        AmtPd = Val(GetFieldValue(Qry))
                        flxComDtls.TextMatrix(jRow, 17) = Format(AmtPd, "#,##0.00")
                        .MoveNext
                        Pbar.Value = Pbar.Value + 1
                        DoEvents
                        pPct = (Pbar.Value / Pbar.Max)
                        lblTotal.Caption = Format$(pPct, "##0.0%")
                        DoEvents
                    Loop
                End If
                .Close
            End With
        End If
    Next i
    flxComDtls.RemoveItem 0
    Pbar.Visible = False
    lblTotal.Caption = ""
    rsCID.Close
End Sub


Private Sub cmdPrnC_Click()
    Dim crAgingCo As New rptAgingCo
    Dim MoYr() As String
    
    If txtSQL.Text = "" Then
        MsgBox "No Record to display!", vbExclamation + vbOKOnly, "Print Preview"
        Exit Sub
    End If
    If InStr(txtSQL.Text, "AGINGAR.AGETYPE)='COMPANY'") = 0 Then
        If cboCPeriod.ListIndex = -1 Or cboCPeriod.Text = "" Then Exit Sub
        MoYr = Split(cboCPeriod.Text, "-")
        txtYr.Text = Trim(MoYr(1))
        txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & "1, " & txtYr.Text))
        txtSQL.Text = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON, AGINGAR.REFNO, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
                      "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='COMPANY')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    End If
    'Display Print Preview
    If rsPrn.State = adStateOpen Then rsPrn.Close
    rsPrn.Open txtSQL.Text, LYRcon, adOpenDynamic, adLockOptimistic
    crAgingCo.DiscardSavedData
    crAgingCo.Database.SetDataSource rsPrn
    crAgingCo.ParameterFields(1).ClearCurrentValueAndRange
    crAgingCo.ParameterFields(1).AddCurrentValue "For the Month of " & cboCPeriod.Text
    crAgingCo.ParameterFields(2).ClearCurrentValueAndRange
    crAgingCo.ParameterFields(2).AddCurrentValue "Leah Flor V. Delima"
    crAgingCo.ParameterFields(3).ClearCurrentValueAndRange
    crAgingCo.ParameterFields(3).AddCurrentValue "Manager"
    crAgingCo.ParameterFields(4).ClearCurrentValueAndRange
    crAgingCo.ParameterFields(4).AddCurrentValue "Lorenzo C. Rubinos Jr."
    crAgingCo.ParameterFields(5).ClearCurrentValueAndRange
    crAgingCo.ParameterFields(5).AddCurrentValue "Area Operation's Manager"
    CRViewer1.ReportSource = crAgingCo
    CRViewer1.ViewReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.EnableExportButton = False
    tabAging.Tab = 3
    lblPrnPrvw.Caption = "AGING OF COMPANY TYPE CUSTOMERS"
End Sub

Private Sub cmdPrnCD_Click()
    Dim crAgingCD As New rptAgingCD
    
    If cboComDtls.ListCount = 0 Then Exit Sub
    If cboComDtls.ListIndex < 1 Then
        txtSQL.Text = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, AGINGAR.REFNO, AGINGAR.REFDATE, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
                      "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='COMPDETAILS')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    Else
        txtSQL.Text = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, AGINGAR.REFNO, AGINGAR.REFDATE, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
                      "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((CUSTOMER.CUSTOMERNAME)='" & cboComDtls.Text & "') AND ((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='COMPDETAILS')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    End If
    'Display Print Preview
    If rsPrn.State = adStateOpen Then rsPrn.Close
    rsPrn.Open txtSQL.Text, LYRcon, adOpenDynamic, adLockOptimistic
    crAgingCD.DiscardSavedData
    crAgingCD.Database.SetDataSource rsPrn
    crAgingCD.ParameterFields(1).ClearCurrentValueAndRange
    crAgingCD.ParameterFields(1).AddCurrentValue "For the Month of " & cboCPeriod.Text
    crAgingCD.ParameterFields(2).ClearCurrentValueAndRange
    crAgingCD.ParameterFields(2).AddCurrentValue "Leah Flor V. Delima"
    crAgingCD.ParameterFields(3).ClearCurrentValueAndRange
    crAgingCD.ParameterFields(3).AddCurrentValue "Manager"
    crAgingCD.ParameterFields(4).ClearCurrentValueAndRange
    crAgingCD.ParameterFields(4).AddCurrentValue "Lorenzo C. Rubinos Jr."
    crAgingCD.ParameterFields(5).ClearCurrentValueAndRange
    crAgingCD.ParameterFields(5).AddCurrentValue "Area Operation's Manager"
    If cboComDtls.Text = "" Then cboComDtls.ListIndex = 0
    crAgingCD.ParameterFields(6).ClearCurrentValueAndRange
    crAgingCD.ParameterFields(6).AddCurrentValue "Detailed Aging of " & cboComDtls.Text
    CRViewer1.ReportSource = crAgingCD
    CRViewer1.ViewReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.EnableExportButton = False
    tabAging.Tab = 3
    lblPrnPrvw.Caption = "Detailed Aging of " & cboComDtls.Text
End Sub

Private Sub cmdPrnI_Click()
    Dim crAgingInd As New rptAgingIN
        
    If txtSQL.Text = "" Then
        MsgBox "No Record to display!", vbExclamation + vbOKOnly, "Print Preview"
        Exit Sub
    End If
    'Display Print Preview
    If rsPrn.State = adStateOpen Then rsPrn.Close
    rsPrn.Open txtSQL.Text, LYRcon, adOpenDynamic, adLockOptimistic
    crAgingInd.DiscardSavedData
    crAgingInd.Database.SetDataSource rsPrn
    crAgingInd.ParameterFields(1).ClearCurrentValueAndRange
    crAgingInd.ParameterFields(1).AddCurrentValue "For the Month of " & cboIPeriod.Text
    crAgingInd.ParameterFields(2).ClearCurrentValueAndRange
    crAgingInd.ParameterFields(2).AddCurrentValue "Leah Flor V. Delima"
    crAgingInd.ParameterFields(3).ClearCurrentValueAndRange
    crAgingInd.ParameterFields(3).AddCurrentValue "Manager"
    crAgingInd.ParameterFields(4).ClearCurrentValueAndRange
    crAgingInd.ParameterFields(4).AddCurrentValue "Lorenzo C. Rubinos Jr."
    crAgingInd.ParameterFields(5).ClearCurrentValueAndRange
    crAgingInd.ParameterFields(5).AddCurrentValue "Area Operation's Manager"
    CRViewer1.ReportSource = crAgingInd
    CRViewer1.ViewReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.EnableExportButton = False
    tabAging.Tab = 3
    lblPrnPrvw.Caption = "AGING OF INDIVIDUAL TYPE CUSTOMERS"
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Unload Me
    MotoRep.tbMain.Visible = True
End Sub

Private Sub flxAgingI_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton And flxAgingI.Row > 0 Then
        Call PopupMenu(mnuIND)
    End If
End Sub

Private Sub Form_Load()
    Dim rsBill As New ADODB.Recordset
    Dim Qry As String, Mo As Integer, Yr As Integer
    
    Qry = "SELECT AgingAR.AGEMONTH, AgingAR.AGEYEAR From AgingAR " & _
          "WHERE (((AgingAR.AGETYPE)='COMPANY')) GROUP BY AgingAR.AGEMONTH, AgingAR.AGEYEAR ORDER BY AgingAR.AGEYEAR DESC , AgingAR.AGEMONTH DESC;"

    With rsBill
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Mo = !AgeMonth
            Yr = !AGEYEAR
            Do While Not .EOF
                cboCPeriod.AddItem MonthName(!AgeMonth) & " - " & !AGEYEAR
                .MoveNext
            Loop
        End If
        .Close
    End With
    tabAging.Tab = 0
    If cboCPeriod.ListCount > 0 Then
        cboCPeriod.ListIndex = 0
    Else
        Qry = "SELECT AGINGAR.CUSTOMERID, AGINGAR.AGEMONTH, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON, AGINGAR.REFNO, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
              "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & Mo & ") AND ((AGINGAR.AGEYEAR)=" & Yr & ") AND ((AGINGAR.AGETYPE)='COMPANY')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    
        rsCAging.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Set flxAgingC.DataSource = rsCAging
        flxAgingC.Refresh
        'Display Company Details
        Qry = "SELECT AGINGAR.CUSTOMERID, AGINGAR.AGEMONTH, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, AGINGAR.REFNO, AGINGAR.REFDATE, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
              "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & Mo & ") AND ((AGINGAR.AGEYEAR)=" & Yr & ") AND ((AGINGAR.AGETYPE)='COMPDETAILS')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
        If rsCDAging.State = adStateOpen Then rsCDAging.Close
        rsCDAging.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Set flxComDtls.DataSource = rsCDAging
        flxComDtls.Refresh
    End If
    MotoRep.sBar.Panels(6).Text = "Process Customer Aging of A/R"
    MotoRep.tbMain.Visible = False
    ActRow = 1
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsCAging.State = adStateOpen Then rsCAging.Close
    If rsIAging.State = adStateOpen Then rsIAging.Close
    If rsCDAging.State = adStateOpen Then rsCDAging.Close
    If rsPrn.State = adStateOpen Then rsPrn.Close
End Sub

Private Sub mnuReProInd_Click()
    Dim rsCust As New ADODB.Recordset
    Dim rsAging As New ADODB.Recordset
    Dim jRow As Integer, CustID As Integer, dPD As Integer, ActAmt As Currency, Bal As Currency, mTerm As Integer, MyDDif As Integer
    Dim Due30 As Currency, Due60 As Currency, Due90 As Currency, Over90 As Currency, TtlOverDue As Currency, Currnt As Currency, AmtPd As Currency, MDue As Currency
    Dim AgeStart As Date, AgeEnd As Date, dTrans As Date, dStart As Date, i As Integer, pPct As Single, InvAge As Integer, MyNP As Integer, DnxtM As Currency
    Dim Qry As String, Tmp As String, MyRef As String, MyTerm As String, MySchm As String, RefType As String
    
    AgeStart = txtDueDate.Text
    AgeEnd = DateAdd("d", -1, DateAdd("m", 1, AgeStart))
    'Start process
    If rsIAging.State = adStateOpen Then rsIAging.Close
    jRow = flxAgingI.Row
    CustID = Val(flxAgingI.TextMatrix(jRow, 1))
    If InStr(flxAgingI.TextMatrix(jRow, 5), "Inv") > 0 Then
        MyRef = Trim(Replace$(flxAgingI.TextMatrix(jRow, 5), "Inv#", ""))
        RefType = "INV"
        Qry = "SELECT ONACCTERMS.* From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & MyRef & "') AND ((ONACCTERMS.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
    ElseIf InStr(flxAgingI.TextMatrix(jRow, 5), "JO") > 0 Then
        MyRef = Trim(Replace$(flxAgingI.TextMatrix(jRow, 5), "JO#", ""))
        RefType = "JO"
        Qry = "SELECT ONACCTERMS.* From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.JOBORDERNO)='" & MyRef & "'));"
    Else
        MyRef = ""
    End If
    If MyRef = "" Then Exit Sub
    ' skip customer without open account
    If CtrlNumExist(Qry) Then
        If RefType = "INV" Then
            Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, ONACCTERMS.SPINVNO, ONACCTERMS.JOBORDERNO, ONACCTERMS.JOBORDERID, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.STATUS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS " & _
                  "FROM CUSTOMER INNER JOIN ONACCTERMS ON CUSTOMER.CUSTOMERID = ONACCTERMS.CUSTOMERID WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.SPINVNO)='" & MyRef & "') AND ((ONACCTERMS.STATUS)='OPEN'));"
            Tmp = "SELECT Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger WHERE (((LEDGER.SPINVNO)='" & MyRef & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ") AND ((LEDGER.JOBID) Is Null Or (LEDGER.JOBID)=0));"
        Else
            Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, ONACCTERMS.SPINVNO, ONACCTERMS.JOBORDERNO, ONACCTERMS.JOBORDERID, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.STATUS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS " & _
                  "FROM CUSTOMER INNER JOIN ONACCTERMS ON CUSTOMER.CUSTOMERID = ONACCTERMS.CUSTOMERID WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.JOBORDERNO)='" & MyRef & "') AND ((ONACCTERMS.STATUS)='OPEN'));"
            Tmp = "SELECT Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger WHERE (((LEDGER.SPINVNO)='" & MyRef & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ") AND ((LEDGER.JOBID)>0));"
        End If
        rsCust.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsCust.RecordCount > 0 Then
            rsCust.MoveFirst
            If Not IsNull(rsCust!Npays) Then
                MyNP = rsCust!Npays
                If Not IsNull(rsCust!PayScheme) Then
                    MySchm = rsCust!PayScheme
                    MyTerm = MyNP & "-" & Left$(rsCust!PayScheme, 1)
                Else
                    MyTerm = MyNP
                    MySchm = ""
                End If
            Else
                MyNP = 0
                If Not IsNull(rsCust!PayScheme) Then
                    MySchm = rsCust!PayScheme
                    MyTerm = Left$(rsCust!PayScheme, 1)
                Else
                    MySchm = ""
                    MyTerm = ""
                End If
            End If
            If IsNull(rsCust!TRANSDATE) Then
                InvAge = 0
            Else
                InvAge = DateDiff("d", rsCust!TRANSDATE, DateSerial(Year(txtDueDate.Text), Month(txtDueDate.Text), Day(rsCust!TRANSDATE)))
            End If
            dTrans = rsCust!TRANSDATE
            ActAmt = IIf(IsNull(rsCust!AcctAmount), 0, rsCust!AcctAmount)
            AmtPd = Val(GetFieldValue(Tmp))
            Tmp = AmtPd
            Bal = IIf(IsNull(rsCust!BALANCE), 0, rsCust!BALANCE)
            If MySchm = "WEEKLY" And MyNP > 0 Then
                MDue = IIf(MyNP >= 4, (ActAmt / MyNP) * 4, ActAmt)
                mTerm = IIf(MyNP >= 4, MyNP \ 4, 1)
                If mTerm > 1 And (MyNP / 4) > mTerm Then mTerm = mTerm + 1
            ElseIf MySchm = "BI-MONTHLY" And MyNP > 0 Then
                MDue = IIf(MyNP >= 2, (ActAmt / MyNP) * 2, ActAmt)
                mTerm = IIf(MyNP >= 2, MyNP \ 2, 1)
                If mTerm > 1 And (MyNP / 2) > mTerm Then mTerm = mTerm + 1
            ElseIf MySchm = "MONTHLY" And MyNP > 0 Then
                MDue = ActAmt / MyNP
                mTerm = MyNP
            End If
            With flxInd
                .Rows = 1
                .Cols = mTerm + 1
                .AddItem "Monthly Due"
                .AddItem "Due Date"
                .AddItem "Aging Sked"
                .AddItem "Balance"
                'initialize value of variables.
                DnxtM = 0: Currnt = 0: Due30 = 0: Due60 = 0: Due90 = 0: Over90 = 0
                dStart = DateSerial(Year(AgeStart), Month(AgeStart), Day(dTrans))
                For i = 1 To .Cols - 1
                    .TextMatrix(1, i) = IIf(ActAmt > MDue, MDue, ActAmt)
                    ActAmt = ActAmt - Val(.TextMatrix(1, i))
                    If i = 1 Then
                        .TextMatrix(2, i) = DateAdd("m", 1, dTrans)
                    Else
                        .TextMatrix(2, i) = DateAdd("m", 1, DateValue(.TextMatrix(2, i - 1)))
                    End If
                    .TextMatrix(4, i) = IIf(AmtPd > Val(.TextMatrix(1, i)), 0, Val(.TextMatrix(1, i)) - AmtPd)
                    If AmtPd > 0 Then AmtPd = AmtPd - (Val(.TextMatrix(1, i)) - Val(.TextMatrix(4, i)))
                    MyDDif = DateDiff("m", DateValue(.TextMatrix(2, i)), dStart)
                    Select Case MyDDif
                        Case Is < 0
                            .TextMatrix(3, i) = "CURRENT"
                            Currnt = Currnt + Val(.TextMatrix(4, i))
                        Case 0
                            .TextMatrix(3, i) = "DUE NEXT MONTH"
                            DnxtM = DnxtM + Val(.TextMatrix(4, i))
                        Case 1
                            .TextMatrix(3, i) = "1 TO 30"
                            Due30 = Due30 + Val(.TextMatrix(4, i))
                        Case 2
                            .TextMatrix(3, i) = "31 TO 60"
                            Due60 = Due60 + Val(.TextMatrix(4, i))
                        Case 3
                            .TextMatrix(3, i) = "61 TO 90"
                            Due90 = Due90 + Val(.TextMatrix(4, i))
                        Case Is > 3
                            .TextMatrix(3, i) = "OVER 90"
                            Over90 = Over90 + Val(.TextMatrix(4, i))
                    End Select
                Next i
            End With
            With flxAgingI
                .TextMatrix(jRow, 7) = MyTerm
                .TextMatrix(jRow, 8) = InvAge
                .TextMatrix(jRow, 9) = Format(DnxtM, "#,##0.00")
                .TextMatrix(jRow, 10) = Format(Currnt, "#,##0.00")
                .TextMatrix(jRow, 11) = Format(Due30, "#,##0.00")
                .TextMatrix(jRow, 12) = Format(Due60, "#,##0.00")
                .TextMatrix(jRow, 13) = Format(Due90, "#,##0.00")
                .TextMatrix(jRow, 14) = Format(Over90, "#,##0.00")
                .TextMatrix(jRow, 15) = Format(Due30 + Due60 + Due90 + Over90, "#,##0.00")
                .TextMatrix(jRow, 16) = Format(Due30 + Due60 + Due90 + Over90 + DnxtM, "#,##0.00")
                .TextMatrix(jRow, 17) = Format(Bal, "#,##0.00")
                AmtPd = Val(Tmp)
                If RefType = "JO" Then
                    Qry = "SELECT Ledger.REFDATE From Ledger WHERE (((Ledger.REFDATE)>=#" & AgeStart & "# And (Ledger.REFDATE)<=#" & AgeEnd & "#) AND ((Ledger.CUSTOMERID)=" & CustID & ") AND ((Ledger.REFTYPE)='Payment') AND ((Ledger.JOBID)>0) AND ((Ledger.SPINVNO)='" & MyRef & "')) ORDER BY Ledger.REFDATE DESC;"
                Else
                    Qry = "SELECT Ledger.REFDATE From Ledger WHERE (((Ledger.REFDATE)>=#" & AgeStart & "# And (Ledger.REFDATE)<=#" & AgeEnd & "#) AND ((Ledger.CUSTOMERID)=" & CustID & ") AND ((Ledger.REFTYPE)='Payment') AND ((Ledger.JOBID)=0 Or (Ledger.JOBID) Is Null) AND ((Ledger.SPINVNO)='" & MyRef & "')) ORDER BY Ledger.REFDATE DESC;"
                End If
                Tmp = GetFieldValue(Qry)
                If Tmp <> "" Then
                    dPD = Day(DateValue(Tmp))
                Else
                    dPD = 0
                End If
                .TextMatrix(jRow, 18) = dPD
                .TextMatrix(jRow, 19) = Format(AmtPd, "#,##0.00")
            End With
        End If
        rsCust.Close
    End If
    With flxAgingI
        Qry = "SELECT AGINGAR.* From AgingAR " & _
              "WHERE (((AGINGAR.CUSTOMERID)=" & CustID & ") AND ((AGINGAR.AGETYPE)='INDIVIDUAL') AND ((AGINGAR.AGEMONTH)=" & Val(txtMo.Text) & ") AND ((AGINGAR.AGEYEAR)=" & Val(txtYr.Text) & ") AND ((AGINGAR.REFNO)='" & .TextMatrix(jRow, 5) & "'));"
        rsAging.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsAging.RecordCount > 0 Then
        rsAging.MoveFirst
        rsAging!CUSTOMERID = .TextMatrix(jRow, 1)
        rsAging!AGETYPE = "INDIVIDUAL"
        rsAging!RefNo = .TextMatrix(jRow, 5)
        rsAging!RefDate = .TextMatrix(jRow, 6)
        rsAging!AgeMonth = txtMo.Text
        rsAging!AGEYEAR = txtYr.Text
        rsAging!Term = MyTerm
        rsAging!InvAge = InvAge
        rsAging!DueNxtMo = Val(Replace(.TextMatrix(jRow, 9), ",", ""))
        rsAging!Current = Val(Replace(.TextMatrix(jRow, 10), ",", ""))
        rsAging!Due1to30 = Val(Replace(.TextMatrix(jRow, 11), ",", ""))
        rsAging!Due31to60 = Val(Replace(.TextMatrix(jRow, 12), ",", ""))
        rsAging!Due61to90 = Val(Replace(.TextMatrix(jRow, 13), ",", ""))
        rsAging!Due91UP = Val(Replace(.TextMatrix(jRow, 14), ",", ""))
        rsAging!TtlOverDue = Val(Replace(.TextMatrix(jRow, 15), ",", ""))
        rsAging!CollNxtMo = Val(Replace(.TextMatrix(jRow, 16), ",", ""))
        rsAging!DateDue = Val(Replace(.TextMatrix(jRow, 17), ",", ""))
        rsAging!DatePd = Val(.TextMatrix(jRow, 16))
        rsAging!AmtPd = Val(Replace(.TextMatrix(jRow, 17), ",", ""))
        rsAging.Update
        End If
        rsAging.Close
    End With
    
    MsgBox "Reprocess of aging completed!", vbInformation + vbOKOnly, "Individual Aging"
    
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, j As Integer
    
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxAgingC        '   1   2           3               4                5            6            7          8           9                10             11                12           13            14         15         16         17
        'If .ColHeader(0) = flexColHeaderOff Then .ColHeader(0) = flexColHeaderOn
        .FormatString = "   |  |^ No. |<    CUSTOMER     |<  ADDRESS   |<     MANAGER  |<REFERENCE|>DUE NEXT MO. |>CURRENT|>1-30 OVERDUE|>31-60 OVERDUE |>61-90 OVERDUE|>91 & UP OVERDUE|>TOTAL OVERDUE|>COLLECTABLE|^DATE DUE|^DATE PAID|^AMOUNT PAID"
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = 300
        For i = 6 To 17
            .ColWidth(i) = (.Width - 300) * 0.1
        Next i
        .ColWidth(3) = (.Width - 300) * 0.2
        .ZOrder
        For i = 1 To .Rows - 1
            .TextMatrix(i, 2) = i 'Format(.TextMatrix(i, 3), "mm/dd/yyyy")
            For j = 7 To 17
                If j <> 15 And j <> 16 Then
                    .TextMatrix(i, j) = Format(.TextMatrix(i, j), "#,##0.00")
                End If
            Next j
        Next i
    End With
    txtFind.Top = cmdFindC.Top
    With flxComDtls        '  1   2           3               4                5            6            7          8           9                10             11                12           13            14         15         16         17
        .ColHeader(0) = flexColHeaderOn
        .FormatString = "   |  |^ No. |<    CUSTOMER     |<  ADDRESS   |<     REF. NO.  |<REF. DATE|>DUE NEXT MO. |>CURRENT|>1-30 OVERDUE|>31-60 OVERDUE |>61-90 OVERDUE|>91 & UP OVERDUE|>TOTAL OVERDUE|>COLLECTABLE|^DATE DUE|^DATE PAID|>AMOUNT PAID"
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = 300
        For i = 6 To 17
            .ColWidth(i) = (.Width - 300) * 0.1
        Next i
        .ColWidth(3) = (.Width - 300) * 0.2
        .ZOrder
        For i = 1 To .Rows - 1
            .TextMatrix(i, 2) = i 'Format(.TextMatrix(i, 3), "mm/dd/yyyy")
            .TextMatrix(i, 6) = Format(.TextMatrix(i, 6), "mm/dd/yyyy")
            For j = 7 To 17
                If j <> 15 And j <> 16 Then
                    .TextMatrix(i, j) = Format(.TextMatrix(i, j), "#,##0.00")
                End If
            Next j
        Next i
    End With
    txtFindCD.Top = cmdFindCD.Top

End Sub


Private Sub tabAging_Click(PreviousTab As Integer)
    Dim MoYr() As String
    
    If tabAging.Tab = PreviousTab Then Exit Sub
    If tabAging.Tab = 0 Then
        If cboCPeriod.ListIndex = -1 Or cboCPeriod.Text = "" Then Exit Sub
        MoYr = Split(cboCPeriod.Text, "-")
        txtYr.Text = Trim(MoYr(1))
        txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & "1, " & txtYr.Text))
        txtSQL.Text = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON, AGINGAR.REFNO, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.DATEDUE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
                  "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='COMPANY')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    ElseIf tabAging.Tab = 2 Then
        If cboIPeriod.ListIndex = -1 Or cboIPeriod.Text = "" Then Exit Sub
        MoYr = Split(cboIPeriod.Text, "-")
        txtYr.Text = Trim(MoYr(1))
        txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & "1, " & txtYr.Text))
        txtDueDate.Text = DateSerial(txtYr.Text, txtMo.Text, 1)
        txtSQL.Text = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, AGINGAR.REFNO, AGINGAR.REFDATE, AGINGAR.TERM, AGINGAR.INVAGE, AGINGAR.REFAMT, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.BALANCE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
                      "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & txtMo.Text & ") AND ((AGINGAR.AGEYEAR)=" & txtYr.Text & ") AND ((AGINGAR.AGETYPE)='INDIVIDUAL')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    End If
End Sub


Private Sub Timer1_Timer()
    Dim rsBill As New ADODB.Recordset
    Dim Qry As String, Mo As Integer, Yr As Integer
    
    If ProInd Then
        Mo = txtMo.Text
        Yr = txtYr.Text
    Else
        Qry = "SELECT AgingAR.AGEMONTH, AgingAR.AGEYEAR From AgingAR " & _
              "WHERE (((AgingAR.AGETYPE)='INDIVIDUAL')) GROUP BY AgingAR.AGEMONTH, AgingAR.AGEYEAR ORDER BY AgingAR.AGEYEAR DESC , AgingAR.AGEMONTH DESC;"
    
        With rsBill
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Mo = !AgeMonth
                Yr = !AGEYEAR
                Do While Not .EOF
                    cboIPeriod.AddItem MonthName(!AgeMonth) & " - " & !AGEYEAR
                    .MoveNext
                Loop
            End If
            .Close
        End With
        If cboIPeriod.ListCount > 0 Then cboIPeriod.ListIndex = 0
    End If
    Qry = "SELECT AGINGAR.CUSTOMERID, AGINGAR.AGEMONTH, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, AGINGAR.REFNO, AGINGAR.REFDATE, AGINGAR.TERM, AGINGAR.INVAGE, AGINGAR.REFAMT, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.BALANCE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
          "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & Mo & ") AND ((AGINGAR.AGEYEAR)=" & Yr & ") AND ((AGINGAR.AGETYPE)='INDIVIDUAL')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    txtSQL.Text = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, AGINGAR.REFNO, AGINGAR.REFDATE, AGINGAR.TERM, AGINGAR.INVAGE, AGINGAR.REFAMT, AGINGAR.DUENXTMO, AGINGAR.CURRENT, AGINGAR.DUE1TO30, AGINGAR.DUE31TO60, AGINGAR.DUE61TO90, AGINGAR.DUE91UP, AGINGAR.TTLOVERDUE, AGINGAR.COLLNXTMO, AGINGAR.BALANCE, AGINGAR.DATEPD, AGINGAR.AMTPD " & _
                  "FROM AGINGAR INNER JOIN CUSTOMER ON AGINGAR.CUSTOMERID = CUSTOMER.CUSTOMERID WHERE (((AGINGAR.AGEMONTH)=" & Mo & ") AND ((AGINGAR.AGEYEAR)=" & Yr & ") AND ((AGINGAR.AGETYPE)='INDIVIDUAL')) ORDER BY CUSTOMER.CUSTOMERNAME, AGINGAR.REFNO;"
    If rsIAging.State = adStateOpen Then rsIAging.Close
    rsIAging.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxAgingI.DataSource = rsIAging
    flxAgingI.Refresh
    With flxAgingI        '  1   2           3               4                5            6           7     8           9           10          11             12           13              14              15            16          17          18          19          20
        'If .ColHeader(0) = flexColHeaderOff Then .ColHeader(0) = flexColHeaderOn
        .FormatString = "   |  |^ No. |<    CUSTOMER     |<  ADDRESS   |<     REF. NO.  |<REF. DATE|^TERM |^AGE  |>REF. AMOUNT|>DUE NEXT MO. |>CURRENT|>1-30 OVERDUE|>31-60 OVERDUE |>61-90 OVERDUE|>91 & UP OVERDUE|>TOTAL OVERDUE|>COLLECTABLE|>BALANCE|^DATE PAID|>AMOUNT PAID"
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = 300
        For i = 6 To 19
            .ColWidth(i) = (.Width - 300) * 0.1
        Next i
        .ColWidth(3) = (.Width - 300) * 0.2
        .ZOrder
        For i = 1 To .Rows - 1
            .TextMatrix(i, 2) = i 'Format(.TextMatrix(i, 3), "mm/dd/yyyy")
            .TextMatrix(i, 6) = Format(.TextMatrix(i, 6), "mm/dd/yyyy")
            For j = 9 To 20
                If j <> 18 Then
                    .TextMatrix(i, j) = Format(.TextMatrix(i, j), "#,##0.00")
                End If
            Next j
        Next i
    End With
    txtFindI.Top = cmdFindI.Top
    ProInd = False
    Timer1.Enabled = False
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 2
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFindc_Click
End Sub

Private Sub txtFindCD_Change()
    fRow = 1
    fCol = 2
End Sub

Private Sub txtFindCD_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFindCD_Click
End Sub

Private Sub txtFindI_Change()
    fRow = 1
    fCol = 2
End Sub

Private Sub txtFindI_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFindI_Click
End Sub
